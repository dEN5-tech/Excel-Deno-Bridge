use std::fs;
use std::path::PathBuf;
use std::process::Command;

fn main() {
    println!("cargo:rerun-if-changed=build.rs");
    println!("cargo:rerun-if-changed=DenoCore.bas.template");
    println!("cargo:rerun-if-changed=UserScripts.bas.template");
    println!("cargo:rerun-if-changed=assemble_xlsm.ps1");

    if std::env::var("CARGO_CFG_TARGET_OS").unwrap_or_default() == "windows" {
        println!("cargo:rustc-cdylib-link-arg=/DEF:exports.def");

        let manifest_dir = PathBuf::from(std::env::var("CARGO_MANIFEST_DIR").unwrap());
        let dll_path = manifest_dir
            .join("target")
            .join("x86_64-pc-windows-msvc")
            .join("release")
            .join("excel_deno_bridge.dll");
        let out_xlsm = manifest_dir.join("dist").join("my_deno_app.xlsm");
        let script_path = manifest_dir.join("assemble_xlsm.ps1");

        let _ = fs::create_dir_all(manifest_dir.join("dist"));

        if std::env::var("GEN_XLSM").is_ok() && dll_path.exists() {
            if let Ok(status) = Command::new("powershell")
                .args([
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    script_path.to_string_lossy().as_ref(),
                    "-dllPath",
                    dll_path.to_string_lossy().as_ref(),
                    "-basDir",
                    manifest_dir.to_string_lossy().as_ref(),
                    "-outPath",
                    out_xlsm.to_string_lossy().as_ref(),
                ])
                .status()
            {
                if status.success() {
                    println!("cargo:warning=>>> XLSM BUNDLE READY at {} <<<", out_xlsm.display());
                }
            }
        }
    }
}