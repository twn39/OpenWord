use libc::c_char;
use std::ffi::{CStr, CString};
use tex2math::{generate_mathml, parse_row, RenderMode};
use winnow::Parser;

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

    let mut cursor = latex_str;
    match parse_row.parse_next(&mut cursor) {
        Ok(ast) => {
            let mathml = generate_mathml(&ast, RenderMode::Display);
            if mathml.is_empty() {
                return std::ptr::null_mut();
            }

            // Wrap in <math> tag
            let full_mathml = format!(
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">{}</math>",
                mathml
            );

            match CString::new(full_mathml) {
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
