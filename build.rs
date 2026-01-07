extern crate winres;

use std::env;
use std::fs::File;
use std::io::{BufRead, BufReader, Write};
use std::path::Path;

fn main() {
    let mut res = winres::WindowsResource::new();
    res.set_resource_file("resources/resource.rc");
    res.compile().unwrap();
    
    parse_resource_h(Path::new("resources/resource.h"), Path::new("src\\resource_consts.rs"));
}

// Copy C style defines to Rust const values
fn parse_resource_h(h_fname: &Path, rs_fname: &Path) {
    let out_dir = env::var("CARGO_MANIFEST_DIR").unwrap();
    let dest_path = Path::new(&out_dir).join(rs_fname);
    let mut f = File::create(&dest_path).unwrap();
    
    let input_file = File::open(h_fname).unwrap();
    let reader = BufReader::new(input_file);

    let mut parsing = false;

    for line in reader.lines() {
        let line = line.unwrap();
        if line.starts_with("// RUST") {
            parsing = true;
            continue;
        }

        if parsing && line.starts_with("#define ") {
            let parts: Vec<_> = line.split_whitespace().collect();
            if parts.len() >= 3 {
                let name = parts[1];
                let value = parts[2];
                writeln!(f, "pub const {}: u16 = {};", name, value).unwrap();
            }
        }
    }
}