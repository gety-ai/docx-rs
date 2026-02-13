use criterion::{criterion_group, criterion_main, Criterion, Throughput, BenchmarkId};
use std::io::Cursor;

mod setup;
use setup::LARGE_DOC_XML;

fn bench_xml_parsing(c: &mut Criterion) {
    let xml_data = &*LARGE_DOC_XML;
    
    let mut group = c.benchmark_group("Large Document Parsing");
    group.throughput(Throughput::Bytes(xml_data.len() as u64));
    group.sample_size(20); // Lower sample size for large files
    group.measurement_time(std::time::Duration::from_secs(15));

    // 1. Current Implementation: xml-rs
    group.bench_function("xml-rs (current)", |b| {
        b.iter(|| {
            let parser = xml::reader::EventReader::new(Cursor::new(xml_data));
            for e in parser {
                let _ = e.unwrap(); // Unwrap to simulate actual checking cost
            }
        })
    });

    // 2. Candidate A: quick-xml (Owned/Buffered)
    // Closest behavior to xml-rs (allocates strings)
    group.bench_function("quick-xml (owned)", |b| {
        b.iter(|| {
            let mut reader = quick_xml::Reader::from_reader(Cursor::new(xml_data));
            reader.config_mut().trim_text(true);
            
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf) {
                    Ok(quick_xml::events::Event::Eof) => break,
                    Ok(_) => buf.clear(),
                    Err(e) => panic!("Error: {:?}", e),
                }
            }
        })
    });

    // 3. Candidate B: quick-xml (Zero-Copy)
    // Maximum performance potential (requires refactoring to string slices)
    group.bench_function("quick-xml (zero-copy)", |b| {
        b.iter(|| {
            let mut reader = quick_xml::Reader::from_bytes(xml_data);
            reader.config_mut().trim_text(true);
            
            loop {
                match reader.read_event() { // Returns borrowed Cow or Bytes
                    Ok(quick_xml::events::Event::Eof) => break,
                    Ok(_) => {},
                    Err(e) => panic!("Error: {:?}", e),
                }
            }
        })
    });

    group.finish();
}

criterion_group!(benches, bench_xml_parsing);
criterion_main!(benches);
