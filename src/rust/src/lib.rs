use ovba::{open_project};

use savvy::savvy;
use savvy::{OwnedStringSexp, OwnedListSexp, Sexp, Result};

use std::fs::{read};

extern crate alloc;
use alloc::collections::BTreeMap;

fn as_sexp(map: BTreeMap<String, String>) -> Result<Sexp> {
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

/// `ovbars` allows to parse the OVBA file `vbaProject.bin` which contains
/// macro variables in Office Open XML files
///
/// `ovbar_out()`: Returns a list of macro sources found in the vba binary.
/// `ovbar_meta()`: Returns a list of additional meta data of directories found
/// in the vba binary. The latter is probably not useful for the user.
/// @name ovba
/// @param name the path to the input file
/// @returns A named list with character strings. The list is sorted alphabetically.
/// @export
#[savvy]
fn ovbar_out(name: &str) -> savvy::Result<savvy::Sexp> {
    let data = read(name)?;
    let project = open_project(data)?;

    let mut map = BTreeMap::new();
    for module in &project.modules {
        let src_code_u8 = project.module_source_raw(&module.name)?;
        let src_code = match String::from_utf8(src_code_u8.clone()) {
            Ok(valid_utf8) => valid_utf8,
            Err(_) => {
                src_code_u8.iter().map(|&b| b as char).collect()
            }
        };

        map.insert(module.name.clone(), src_code);
    }

    let list = as_sexp(map)?;

    Ok(list)
}

/// @rdname ovba
/// @export
#[savvy]
fn ovbar_meta(name: &str) -> savvy::Result<savvy::Sexp> {
    // Read raw project container
    let data = read(name)?;
    let project = open_project(data)?;


    let mut map = BTreeMap::new();
    // Iterate over CFB entries
    for (nm, path) in project.list()? {
        map.insert(nm, path);
    }

    let list = as_sexp(map)?;

    Ok(list)
}
