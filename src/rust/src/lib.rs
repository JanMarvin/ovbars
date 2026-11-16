use ovba::open_project;

use savvy::savvy;
use savvy::{OwnedStringSexp, OwnedListSexp, Sexp, Result};

use std::collections::BTreeMap;
use std::fs::read;

fn as_sexp(map: BTreeMap<String, String>) -> Result<Sexp> {
    let mut list = OwnedListSexp::new(map.len(), true)?;

    for (index, (key, value)) in map.iter().enumerate() {
        let mut str = OwnedStringSexp::new(1)?;
        str.set_elt(0, value)?;
        list.set_name_and_value(index, key, str)?;
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
        let src_code = project.module_source(&module.name)?;

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
