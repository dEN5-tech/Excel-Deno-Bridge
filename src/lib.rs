use deno_core::{v8, FastString, JsRuntime, PollEventLoopOptions, RuntimeOptions};
use once_cell::sync::Lazy;
use std::cell::{Cell, RefCell};
use std::collections::HashMap;
use std::mem::ManuallyDrop;
use std::panic::{self, AssertUnwindSafe};
use std::path::PathBuf;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::Mutex;
use std::time::Duration;
use tokio::runtime::Runtime;
use widestring::U16CStr;
use windows::core::{w, BSTR, ComInterface, HSTRING, PCWSTR};
use windows::Win32::Foundation::{HWND, VARIANT_BOOL};
use windows::Win32::System::Com::{
    CoInitializeEx, IDispatch, COINIT_APARTMENTTHREADED, DISPATCH_METHOD, DISPATCH_PROPERTYGET,
    DISPATCH_PROPERTYPUT, DISPPARAMS,
};
use windows::Win32::System::Com::CLSIDFromProgID;
use windows::Win32::System::Ole::{GetActiveObject, DISPID_PROPERTYPUT};
use windows::Win32::System::Variant::{
    VariantClear, VariantInit, VARENUM, VARIANT, VARIANT_0, VARIANT_0_0, VARIANT_0_0_0,
    VT_BOOL, VT_BSTR, VT_DISPATCH, VT_EMPTY, VT_I4, VT_NULL, VT_R8,
};
use windows::Win32::UI::WindowsAndMessaging::{KillTimer, SetTimer};

deno_core::extension!(
    vba_bridge,
    ops = [
        op_read_workbook_file,
        op_set_workbook_dir,
        op_fetch_sync,
        op_excel_proxy_request,
        op_start_event_monitor,
        op_poll_excel_event
    ],
);

static WORKBOOK_DIR: Lazy<Mutex<Option<PathBuf>>> = Lazy::new(|| Mutex::new(None));
static EVENT_QUEUE: Lazy<Mutex<Vec<(String, String)>>> = Lazy::new(|| Mutex::new(Vec::new()));
static EVENT_MONITOR_STARTED: AtomicBool = AtomicBool::new(false);
static NATIVE_TIMER_ID: Lazy<Mutex<usize>> = Lazy::new(|| Mutex::new(0));

#[deno_core::op2]
#[string]
fn op_read_workbook_file(#[string] filename: String) -> Result<String, deno_core::anyhow::Error> {
    let dir_lock = WORKBOOK_DIR
        .lock()
        .map_err(|_| deno_core::anyhow::anyhow!("Workbook path lock poisoned"))?;
    let dir = dir_lock
        .as_ref()
        .ok_or_else(|| deno_core::anyhow::anyhow!("Workbook path not set"))?;

    let content = std::fs::read_to_string(dir.join(filename))?;
    Ok(content)
}

#[deno_core::op2(fast)]
fn op_set_workbook_dir(#[string] dir: String) {
    if let Ok(mut lock) = WORKBOOK_DIR.lock() {
        *lock = Some(PathBuf::from(dir));
    }
}

#[deno_core::op2]
#[string]
fn op_fetch_sync(#[string] url: String) -> Result<String, deno_core::anyhow::Error> {
    let response = reqwest::blocking::get(url)?.text()?;
    Ok(response)
}

#[deno_core::op2(fast)]
fn op_start_event_monitor() {
    if EVENT_MONITOR_STARTED
        .compare_exchange(false, true, Ordering::SeqCst, Ordering::SeqCst)
        .is_err()
    {
        return;
    }

    std::thread::spawn(|| unsafe {
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
        let mut last_seen = String::new();

        loop {
            match get_active_excel_dispatch() {
                Ok(app) => {
                    if let Ok(active_cell) = get_property_dispatch(&app, "ActiveCell") {
                        if let Ok(mut value) = get_property_variant(&active_cell, "Value") {
                            let current = variant_to_plain_string(&value);
                            let _ = VariantClear(&mut value);

                            if current != last_seen {
                                last_seen = current.clone();
                                if let Ok(mut q) = EVENT_QUEUE.lock() {
                                    q.push(("cell_change".to_string(), current));
                                    if q.len() > 128 {
                                        let trim = q.len() - 128;
                                        let _ = q.drain(..trim);
                                    }
                                }
                            }
                        }
                    }
                    std::thread::sleep(Duration::from_millis(150));
                }
                Err(_e) => {
                    // Excel can be temporarily busy (editing cell/modal dialogs).
                    // Back off and retry.
                    std::thread::sleep(Duration::from_millis(300));
                }
            }
        }
    });
}

#[deno_core::op2]
#[serde]
fn op_poll_excel_event() -> serde_json::Value {
    if let Ok(mut q) = EVENT_QUEUE.lock() {
        if !q.is_empty() {
            let (name, data) = q.remove(0);
            return serde_json::json!({"name":name,"data":data});
        }
    }
    serde_json::Value::Null
}

unsafe fn get_active_excel_dispatch() -> Result<IDispatch, deno_core::anyhow::Error> {
    let clsid = CLSIDFromProgID(w!("Excel.Application"))?;
    let mut unknown = None;
    GetActiveObject(&clsid, None, &mut unknown)?;
    let unknown = unknown.ok_or_else(|| deno_core::anyhow::anyhow!("Excel instance not found"))?;
    Ok(unknown.cast::<IDispatch>()?)
}

unsafe fn get_property_variant(
    dispatch: &IDispatch,
    name: &str,
) -> Result<VARIANT, deno_core::anyhow::Error> {
    let mut disp_id = 0i32;
    let name_h = HSTRING::from(name);
    let name_ptr = PCWSTR(name_h.as_ptr());
    dispatch.GetIDsOfNames(&Default::default(), &name_ptr, 1, 0x0409, &mut disp_id)?;

    let params = DISPPARAMS::default();
    let mut result = VariantInit();
    dispatch.Invoke(
        disp_id,
        &Default::default(),
        0x0409,
        DISPATCH_PROPERTYGET,
        &params,
        Some(&mut result),
        None,
        None,
    )?;
    Ok(result)
}

unsafe fn get_property_dispatch(
    dispatch: &IDispatch,
    name: &str,
) -> Result<IDispatch, deno_core::anyhow::Error> {
    let mut v = get_property_variant(dispatch, name)?;
    let out = if v.Anonymous.Anonymous.vt == VT_DISPATCH {
        v.Anonymous.Anonymous.Anonymous
            .pdispVal
            .as_ref()
            .cloned()
            .ok_or_else(|| deno_core::anyhow::anyhow!("Dispatch property empty"))?
    } else {
        let _ = VariantClear(&mut v);
        return Err(deno_core::anyhow::anyhow!("Property is not dispatch"));
    };
    let _ = VariantClear(&mut v);
    Ok(out)
}

unsafe fn variant_to_plain_string(v: &VARIANT) -> String {
    let vt: VARENUM = v.Anonymous.Anonymous.vt;
    if vt == VT_BSTR {
        return v.Anonymous.Anonymous.Anonymous.bstrVal.to_string();
    }
    if vt == VT_I4 {
        return v.Anonymous.Anonymous.Anonymous.lVal.to_string();
    }
    if vt == VT_R8 {
        return v.Anonymous.Anonymous.Anonymous.dblVal.to_string();
    }
    if vt == VT_BOOL {
        return (v.Anonymous.Anonymous.Anonymous.boolVal.0 != 0).to_string();
    }
    if vt == VT_EMPTY || vt == VT_NULL {
        return String::new();
    }
    "[variant]".to_string()
}

#[deno_core::op2]
#[serde]
fn op_excel_proxy_request(
    #[serde] req: serde_json::Value,
) -> Result<serde_json::Value, deno_core::anyhow::Error> {
    let obj_id = req["id"].as_u64().unwrap_or(0) as u32;
    let member = req["member"].as_str().unwrap_or("");
    let action = req["action"].as_str().unwrap_or("get");
    let args = req
        .get("args")
        .and_then(|v| v.as_array())
        .cloned()
        .unwrap_or_default();

    unsafe {
        let dispatch: IDispatch = if obj_id == 0 {
            let clsid = CLSIDFromProgID(w!("Excel.Application"))?;
            let mut unknown = None;
            GetActiveObject(&clsid, None, &mut unknown)?;
            let unknown =
                unknown.ok_or_else(|| deno_core::anyhow::anyhow!("Excel instance not found"))?;
            unknown.cast::<IDispatch>()?
        } else {
            COM_OBJECTS.with(|m| {
                m.borrow()
                    .get(&obj_id)
                    .cloned()
                    .ok_or_else(|| deno_core::anyhow::anyhow!("Object lost"))
            })?
        };

        let mut disp_id = 0i32;
        let name_h = HSTRING::from(member);
        let name_ptr = PCWSTR(name_h.as_ptr());
        dispatch.GetIDsOfNames(
            &Default::default(),
            &name_ptr,
            1,
            0x0409,
            &mut disp_id,
        )?;

        let mut params = DISPPARAMS::default();
        let mut flags = DISPATCH_PROPERTYGET;
        let mut named_dispid = DISPID_PROPERTYPUT;
        let mut arg_variants: Vec<VARIANT> = Vec::new();

        if action == "set" {
            flags = DISPATCH_PROPERTYPUT;
            let value = req.get("value").cloned().unwrap_or(serde_json::Value::Null);
            arg_variants.push(json_to_variant(&value));
            params.cArgs = 1;
            params.rgvarg = arg_variants.as_mut_ptr();
            params.cNamedArgs = 1;
            params.rgdispidNamedArgs = &mut named_dispid;
        } else if action == "call" {
            flags = DISPATCH_METHOD | DISPATCH_PROPERTYGET;
            if !args.is_empty() {
                for arg in args.iter().rev() {
                    arg_variants.push(json_to_variant(arg));
                }
                params.cArgs = arg_variants.len() as u32;
                params.rgvarg = arg_variants.as_mut_ptr();
            }
        }

        let mut result = VariantInit();
        dispatch.Invoke(
            disp_id,
            &Default::default(),
            0x0409,
            flags,
            &params,
            Some(&mut result),
            None,
            None,
        )?;

        let out = variant_to_json(&result);

        for v in arg_variants.iter_mut() {
            let _ = VariantClear(v);
        }
        let _ = VariantClear(&mut result);

        Ok(out)
    }
}

fn json_to_variant(v: &serde_json::Value) -> VARIANT {
    match v {
        serde_json::Value::String(s) => VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: VT_BSTR,
                    wReserved1: 0,
                    wReserved2: 0,
                    wReserved3: 0,
                    Anonymous: VARIANT_0_0_0 {
                        bstrVal: ManuallyDrop::new(BSTR::from(s.as_str())),
                    },
                }),
            },
        },
        serde_json::Value::Number(n) => {
            if let Some(i) = n.as_i64() {
                VARIANT {
                    Anonymous: VARIANT_0 {
                        Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                            vt: VT_I4,
                            wReserved1: 0,
                            wReserved2: 0,
                            wReserved3: 0,
                            Anonymous: VARIANT_0_0_0 { lVal: i as i32 },
                        }),
                    },
                }
            } else if let Some(f) = n.as_f64() {
                VARIANT {
                    Anonymous: VARIANT_0 {
                        Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                            vt: VT_R8,
                            wReserved1: 0,
                            wReserved2: 0,
                            wReserved3: 0,
                            Anonymous: VARIANT_0_0_0 { dblVal: f },
                        }),
                    },
                }
            } else {
                unsafe { VariantInit() }
            }
        }
        serde_json::Value::Bool(b) => VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: VT_BOOL,
                    wReserved1: 0,
                    wReserved2: 0,
                    wReserved3: 0,
                    Anonymous: VARIANT_0_0_0 {
                        boolVal: VARIANT_BOOL(if *b { -1 } else { 0 }),
                    },
                }),
            },
        },
        serde_json::Value::Null => VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: VT_NULL,
                    wReserved1: 0,
                    wReserved2: 0,
                    wReserved3: 0,
                    Anonymous: VARIANT_0_0_0 { lVal: 0 },
                }),
            },
        },
        _ => VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: VT_EMPTY,
                    wReserved1: 0,
                    wReserved2: 0,
                    wReserved3: 0,
                    Anonymous: VARIANT_0_0_0 { lVal: 0 },
                }),
            },
        },
    }
}

fn variant_to_json(v: &VARIANT) -> serde_json::Value {
    unsafe {
        let vt: VARENUM = v.Anonymous.Anonymous.vt;
        if vt == VT_DISPATCH {
            let maybe_disp = &v.Anonymous.Anonymous.Anonymous.pdispVal;
            if let Some(next_disp) = maybe_disp.as_ref() {
                let new_id = NEXT_OBJ_ID.with(|id| {
                    let current = id.get();
                    id.set(current + 1);
                    current
                });
                COM_OBJECTS.with(|m| {
                    m.borrow_mut().insert(new_id, next_disp.clone());
                });
                return serde_json::json!({"type":"obj","id":new_id});
            }
        }

        if vt == VT_BSTR {
            let b = &v.Anonymous.Anonymous.Anonymous.bstrVal;
            let s = b.to_string();
            return serde_json::json!({"type":"val","value": s});
        }
        if vt == VT_I4 {
            return serde_json::json!({"type":"val","value": v.Anonymous.Anonymous.Anonymous.lVal});
        }
        if vt == VT_R8 {
            return serde_json::json!({"type":"val","value": v.Anonymous.Anonymous.Anonymous.dblVal});
        }
        if vt == VT_BOOL {
            let b = v.Anonymous.Anonymous.Anonymous.boolVal.0 != 0;
            return serde_json::json!({"type":"val","value": b});
        }
        if vt == VT_EMPTY || vt == VT_NULL {
            return serde_json::json!({"type":"val","value": serde_json::Value::Null});
        }

        serde_json::json!({"type":"val","value":"[unsupported variant]"})
    }
}

struct DenoBridge {
    js_runtime: JsRuntime,
    tokio_rt: Runtime,
}

// VBA calls happen on a single STA thread. We initialize once and keep runtime alive
// for the lifetime of the DLL in that thread.
thread_local! {
    static COM_OBJECTS: RefCell<HashMap<u32, IDispatch>> = RefCell::new(HashMap::new());
    static NEXT_OBJ_ID: Cell<u32> = const { Cell::new(1) };

    static BRIDGE: Lazy<RefCell<Option<DenoBridge>>> = Lazy::new(|| {
        unsafe {
            let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
        }
        RefCell::new(init_bridge())
    });
}

fn init_bridge() -> Option<DenoBridge> {
    let tokio_rt = tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
        .ok()?;

    let mut js_runtime = JsRuntime::new(RuntimeOptions {
        extensions: vec![vba_bridge::init_ops()],
        ..Default::default()
    });

    js_runtime
        .execute_script(
            "<init_include>",
            FastString::from(
                "globalThis.bridgeVersion = 'excel_deno_bridge/1.1.0';\n\
                globalThis.onExcelEvent = globalThis.onExcelEvent || null;\n\
                globalThis.__excelEventTimer = globalThis.__excelEventTimer || null;\n\
                globalThis.__runPoll = () => {\n\
                    const ev = Deno.core.ops.op_poll_excel_event();\n\
                    if (ev && globalThis.onExcelEvent) {\n\
                        try { globalThis.onExcelEvent(ev.name, ev.data); } catch (_) {}\n\
                    }\n\
                    return ev || null;\n\
                };\n\
                globalThis.startExcelEvents = () => {\n\
                    Deno.core.ops.op_start_event_monitor();\n\
                    if (!globalThis.__excelEventTimer) {\n\
                        globalThis.__excelEventTimer = setInterval(() => {\n\
                            globalThis.__runPoll();\n\
                        }, 120);\n\
                    }\n\
                    return 'ok';\n\
                };\n\
                const __coerce = (res) => (res && res.type === 'obj') ? __excelObject(res.id) : (res ? res.value : undefined);\n\
                const __excelMember = (objId, member) => new Proxy(function(...args) {}, {\n\
                  get(_t, prop) {\n\
                    if (prop === 'then') return undefined;\n\
                    const base = Deno.core.ops.op_excel_proxy_request({ id: objId, member, action: 'get' });\n\
                    const next = __coerce(base);\n\
                    return (next && next.__isExcelObject) ? next[prop] : next;\n\
                  },\n\
                  set(_t, prop, value) {\n\
                    const base = Deno.core.ops.op_excel_proxy_request({ id: objId, member, action: 'get' });\n\
                    const next = __coerce(base);\n\
                    if (next && next.__isExcelObject) { next[prop] = value; return true; }\n\
                    return false;\n\
                  },\n\
                  apply(_t, _thisArg, args) {\n\
                    const res = Deno.core.ops.op_excel_proxy_request({ id: objId, member, action: 'call', args });\n\
                    return __coerce(res);\n\
                  }\n\
                });\n\
                const __excelObject = (id) => new Proxy(function(){}, {\n\
                  get(_t, prop) {\n\
                    if (prop === 'then') return undefined;\n\
                    if (prop === '__isExcelObject') return true;\n\
                    return __excelMember(id, String(prop));\n\
                  },\n\
                  set(_t, prop, value) {\n\
                    Deno.core.ops.op_excel_proxy_request({ id, member: String(prop), action: 'set', value });\n\
                    return true;\n\
                  }\n\
                });\n\
                globalThis.Excel = __excelObject(0);\n\
                globalThis.__set_workbook_dir = (dir) => {\n\
                    Deno.core.ops.op_set_workbook_dir(String(dir));\n\
                    return 'ok';\n\
                };\n\
                globalThis.fetchSync = (url) => {\n\
                    return Deno.core.ops.op_fetch_sync(String(url));\n\
                };\n\
                globalThis.fetch = async (url) => {\n\
                    const __raw = Deno.core.ops.op_fetch_sync(String(url));\n\
                    return {\n\
                        text: async () => __raw,\n\
                        json: async () => JSON.parse(__raw),\n\
                    };\n\
                };\n\
                globalThis.include = (filename) => {\n\
                    const code = Deno.core.ops.op_read_workbook_file(String(filename));\n\
                    return (0, eval)(code);\n\
                };"
                .to_string(),
            ),
        )
        .ok()?;

    Some(DenoBridge {
        js_runtime,
        tokio_rt,
    })
}

fn to_bstr_ptr(s: &str) -> *mut u16 {
    BSTR::from(s).into_raw().cast_mut()
}

fn pump_bridge_once() {
    BRIDGE.with(|cell| {
        let Ok(mut bridge_opt) = cell.try_borrow_mut() else {
            return;
        };
        let Some(bridge) = bridge_opt.as_mut() else {
            return;
        };

        let _ = bridge.tokio_rt.block_on(async {
            let _ = bridge.js_runtime.execute_script(
                "<heartbeat>",
                FastString::from("if (globalThis.__runPoll) { globalThis.__runPoll(); }".to_string()),
            );
            bridge
                .js_runtime
                .run_event_loop(PollEventLoopOptions::default())
                .await
        });
    });
}

unsafe extern "system" fn native_timer_proc(_hwnd: HWND, _msg: u32, _id: usize, _time: u32) {
    pump_bridge_once();
}

fn execute_js_inner(code_ptr: *const u16) -> *mut u16 {
    if code_ptr.is_null() {
        return to_bstr_ptr("Error: Null Pointer");
    }

    let input_u16 = unsafe { U16CStr::from_ptr_str(code_ptr) };
    let js_code = input_u16.to_string_lossy();

    BRIDGE.with(|cell| {
        let mut bridge_opt = cell.borrow_mut();
        let Some(bridge) = bridge_opt.as_mut() else {
            return to_bstr_ptr("Error: Failed to init Deno");
        };

        let outcome: Result<String, String> = bridge.tokio_rt.block_on(async {
            let user_code_escaped = serde_json::to_string(&js_code)
                .unwrap_or_else(|_| "\"\"".to_string());

            let wrapped_script = format!(
                r#"(function() {{
  try {{
    if (typeof globalThis.fetchSync !== 'function') {{
      globalThis.fetchSync = (url) => Deno.core.ops.op_fetch_sync(String(url));
    }}
    if (typeof globalThis.fetch !== 'function') {{
      globalThis.fetch = async (url) => {{
        const __raw = Deno.core.ops.op_fetch_sync(String(url));
        return {{
          text: async () => __raw,
          json: async () => JSON.parse(__raw),
        }};
      }};
    }}
    if (typeof globalThis.include !== 'function') {{
      globalThis.include = (filename) => {{
        const code = Deno.core.ops.op_read_workbook_file(String(filename));
        return (0, eval)(code);
      }};
    }}

    const __src = {src};
    let __res;

    try {{
      __res = eval(__src);
    }} catch (inner) {{
      // Convenience fallback: allow inline template-style text without backticks.
      if (__src.includes('${{')) {{
        const __safe = __src.replace(/`/g, '\\`');
        __res = eval('`' + __safe + '`');
      }} else {{
        throw inner;
      }}
    }}

    if (__res === undefined) return 'undefined';
    if (__res === null) return 'null';

    if (typeof __res === 'object') {{
      try {{
        return JSON.stringify(__res, null, 2);
      }} catch (_) {{
        return String(__res);
      }}
    }}

    return String(__res);
  }} catch (e) {{
    return 'JS Error: ' + (e && e.message ? e.message : String(e));
  }}
}})()"#,
                src = user_code_escaped
            );

            let value = bridge
                .js_runtime
                .execute_script("<excel_bridge>", FastString::from(wrapped_script))
                .map_err(|e| format!("JS Error: {e}"))?;

            // Heartbeat: advance timers/microtasks so setInterval-based polling keeps working
            bridge
                .js_runtime
                .run_event_loop(PollEventLoopOptions::default())
                .await
                .map_err(|e| format!("Event loop error: {e}"))?;

            let scope = &mut bridge.js_runtime.handle_scope();
            let local = v8::Local::new(scope, value);
            let string_value = local
                .to_string(scope)
                .map(|s| s.to_rust_string_lossy(scope))
                .unwrap_or_else(|| "undefined".to_string());

            Ok(string_value)
        });

        match outcome {
            Ok(out) => to_bstr_ptr(&out),
            Err(err) => to_bstr_ptr(&err),
        }
    })
}

#[no_mangle]
pub extern "system" fn start_native_heartbeat(interval_ms: u32) {
    let interval = if interval_ms == 0 { 100 } else { interval_ms };
    if let Ok(mut id_lock) = NATIVE_TIMER_ID.lock() {
        if *id_lock == 0 {
            unsafe {
                *id_lock = SetTimer(None, 0, interval, Some(native_timer_proc));
            
            }
        }
    }
}

#[no_mangle]
pub extern "system" fn stop_native_heartbeat() {
    if let Ok(mut id_lock) = NATIVE_TIMER_ID.lock() {
        if *id_lock != 0 {
            unsafe {
                let _ = KillTimer(None, *id_lock);
            }
            *id_lock = 0;
        }
    }
}

#[no_mangle]
pub extern "system" fn execute_js(code_ptr: *const u16) -> *mut u16 {
    match panic::catch_unwind(AssertUnwindSafe(|| execute_js_inner(code_ptr))) {
        Ok(ptr) => ptr,
        Err(_) => to_bstr_ptr("Error: Rust Panic"),
    }
}

#[no_mangle]
pub extern "system" fn free_bstr(ptr: *mut u16) {
    if ptr.is_null() {
        return;
    }

    unsafe {
        // Rebuild owned BSTR and let Drop call SysFreeString exactly once.
        let _owned = BSTR::from_raw(ptr);
    }
}

#[no_mangle]
pub extern "system" fn set_workbook_path(path_ptr: *const u16) {
    if path_ptr.is_null() {
        return;
    }

    let path = unsafe { U16CStr::from_ptr_str(path_ptr) }.to_string_lossy();
    if let Ok(mut lock) = WORKBOOK_DIR.lock() {
        *lock = Some(PathBuf::from(path));
    }
}
