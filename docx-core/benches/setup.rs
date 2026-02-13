use std::fs::File;
use std::io::Read;
use std::path::Path;
use once_cell::sync::Lazy;

// Lazy static to hold the large XML content in memory
pub static LARGE_DOC_XML: Lazy<Vec<u8>> = Lazy::new(|| {
    let path = Path::new("../fixtures/large_file/The Routledge Handbook of Translation and Philosophy.docx");
    let file = File::open(path).expect("Test fixture not found. Ensure you are running benchmarks from docx-core/ or root.");
    let mut archive = zip::ZipArchive::new(file).expect("Failed to open zip");
    
    // Extract purely the document.xml content
    let mut xml_content = Vec::new();
    archive
        .by_name("word/document.xml")
        .expect("document.xml missing")
        .read_to_end(&mut xml_content)
        .expect("Failed to read xml");
        
    xml_content
});
