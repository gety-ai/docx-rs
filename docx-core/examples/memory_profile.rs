use dhat;

#[global_allocator]
static ALLOC: dhat::Alloc = dhat::Alloc;

// Copy the setup logic or import it if you make setup.rs a library module
// For simplicity, we assume we read the file here directly
fn get_test_data() -> Vec<u8> {
    use std::io::Read;
    use std::path::Path;
    let path = Path::new("../fixtures/large_file/The Routledge Handbook of Translation and Philosophy.docx");
    let file = std::fs::File::open(path).expect("Test fixture not found");
    let mut archive = zip::ZipArchive::new(file).unwrap();
    let mut buf = Vec::new();
    archive.by_name("word/document.xml").unwrap().read_to_end(&mut buf).unwrap();
    buf
}

fn main() {
    let _profiler = dhat::Profiler::new_heap();
    
    println!("Preparing data...");
    let data = get_test_data();
    
    println!("Profiling xml-rs...");
    // Isolate scope for profiling
    {
        let parser = xml::reader::EventReader::new(std::io::Cursor::new(&data));
        let mut count = 0;
        for _ in parser { count += 1; }
        println!("xml-rs events: {}", count);
    }
    
    // NOTE: To profile quick-xml, swap the code block above or create two separate binaries.
    // DHAT works best when profiling one distinct workload per run.
}
