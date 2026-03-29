use std::ffi::{CStr, CString};
use libc::c_char;
use math_core::{LatexToMathML, MathCoreConfig, MathDisplay};

#[no_mangle]
pub extern "C" fn latex_to_mathml_c(latex: *const c_char) -> *mut c_char {
    if latex.is_null() { return std::ptr::null_mut(); }
    let c_str = unsafe { CStr::from_ptr(latex) };
    let latex_str = match c_str.to_str() {
        Ok(s) => s,
        Err(_) => return std::ptr::null_mut(),
    };
    let config = MathCoreConfig::default();
    let converter = match LatexToMathML::new(config) {
        Ok(c) => c,
        Err(_) => return std::ptr::null_mut(),
    };
    match converter.convert_with_local_counter(latex_str, MathDisplay::Block) {
        Ok(mathml) => {
            if mathml.is_empty() { return std::ptr::null_mut(); }
            match CString::new(mathml) {
                Ok(c_string) => c_string.into_raw(),
                Err(_) => std::ptr::null_mut(),
            }
        }
        Err(_) => std::ptr::null_mut(),
    }
}

#[no_mangle]
pub extern "C" fn free_rust_string(s: *mut c_char) {
    if !s.is_null() {
        unsafe { drop(CString::from_raw(s)); }
    }
}
