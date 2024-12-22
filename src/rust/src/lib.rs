use ovba::{open_project};

use savvy::savvy;
use savvy::{OwnedStringSexp, OwnedListSexp, Sexp, Result};

use std::fs::{read}; // write
use std::collections::HashMap;

fn as_sexp(map: HashMap<String, String>) -> Result<Sexp> {
    let mut list = OwnedListSexp::new(map.len(), true)?;

    // Convert HashMap to a vector of tuples and sort by the key (String)
    let mut sorted_entries: Vec<_> = map.iter().collect();
    sorted_entries.sort_by(|a, b| a.0.cmp(b.0)); // Sort by key (String)

    for (index, (key, value)) in sorted_entries.iter().enumerate() {
        let mut str = OwnedStringSexp::new(1)?;
        str.set_elt(0, value)?;
        list.set_name_and_value(index, key, str)?; // Correct method call
    }

    Ok(list.into())
}

/// Returns a list of macro sources found in the vba binary.
/// The list is sorted alphabetically.
/// @param name the path to the input file
/// @export
#[savvy]
fn ovbar_out(name: &str) -> savvy::Result<savvy::Sexp> {
    let data = read(name)?;
    let project = open_project(data)?;

    let mut map = HashMap::new();
    for module in &project.modules {
        let src_code_u8 = project.module_source_raw(&module.name)?;
        let src_code = String::from_utf8(src_code_u8.clone()).unwrap_or_else(|_| String::from("Invalid UTF-8"));

        map.insert(module.name.clone(), src_code);
    }

    let list = as_sexp(map)?;

    Ok(list)
}

/// Returns a list of additional meta data of directories found in the vba
/// binary. This is probably not useful for the user.
/// The list is sorted alphabetically.
/// @param name the path to the input file
/// @export
#[savvy]
fn ovbar_meta(name: &str) -> savvy::Result<savvy::Sexp> {
    // Read raw project container
    let data = read(name)?;
    let project = open_project(data)?;


    let mut map = HashMap::new();
    // Iterate over CFB entries
    for (nm, path) in project.list()? {
        map.insert(nm, path);
    }

    let list = as_sexp(map)?;

    Ok(list)
}
