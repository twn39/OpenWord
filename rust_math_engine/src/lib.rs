use std::ffi::{CStr, CString};
use libc::c_char;
use latex2mathml::{latex_to_mathml, DisplayStyle};

#[no_mangle]
pub extern "C" fn latex_to_mathml_c(latex: *const c_char) -> *mut c_char {
    if latex.is_null() {
        return std::ptr::null_mut();
    }

    let c_str = unsafe { CStr::from_ptr(latex) };
    let latex_str = match c_str.to_str() {
        Ok(s) => s,
        Err(_) => return std::ptr::null_mut(),
    };

    // 将 LaTeX 转换为 MathML (Block 模式)
    match latex_to_mathml(latex_str, DisplayStyle::Block) {
        Ok(mathml) => {
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
        unsafe {
            drop(CString::from_raw(s));
        }
    }
}
