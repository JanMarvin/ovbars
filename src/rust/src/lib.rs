use extendr_api::prelude::*;

use ovba::{open_project, Result};
use std::fs::{read}; // write

use std::collections::HashMap;

/// Returns a list of macro sources found in the vba binary.
/// The list is sorted alphabetically.
/// @param name the path to the input file
/// @export
#[extendr]
fn ovbar_out(name: &str) -> Result<List> {
    let data = read(name)?;
    let project = open_project(data)?;


    let mut map = HashMap::new();
    for module in &project.modules {
        let src_code_u8 = project.module_source_raw(&module.name)?;

        let src_code = String::from_utf8(src_code_u8.clone()).unwrap_or_else(|_| String::from("Invalid UTF-8"));

        map.insert(module.name.clone(), src_code);
    }

    let mut sorted_keys: Vec<String> = map.keys().cloned().collect();
    sorted_keys.sort();

    // Create a vector of sorted key-value pairs
    let r_list_elements: Vec<(String, Robj)> = sorted_keys
        .into_iter()
        .map(|key| (key.clone(), Robj::from(map[&key].clone())))
        .collect();

    let r_list = List::from_pairs(r_list_elements);

    // Return the list to R
    Ok(r_list)
}

/// Returns a list of additional meta data of directories found in the vba
/// binary. This is probably not useful for the user.
/// The list is sorted alphabetically.
/// @param name the path to the input file
/// @export
#[extendr]
fn ovbar_meta(name: &str) -> Result<List> {
    // Read raw project container
    let data = read(name)?;
    let project = open_project(data)?;

    let mut map = HashMap::new();
    // Iterate over CFB entries
    for (nm, path) in project.list()? {
        map.insert(nm, path);
    }

    let mut sorted_keys: Vec<String> = map.keys().cloned().collect();
    sorted_keys.sort();

    // Create a vector of sorted key-value pairs
    let r_list_elements: Vec<(String, Robj)> = sorted_keys
        .into_iter()
        .map(|key| (key.clone(), Robj::from(map[&key].clone())))
        .collect();

    let r_list = List::from_pairs(r_list_elements);

    Ok(r_list)
}

// Macro to generate exports.
// This ensures exported functions are registered with R.
// See corresponding C code in `entrypoint.c`.
extendr_module! {
    mod ovbars;
    fn ovbar_out;
    fn ovbar_meta;
}
