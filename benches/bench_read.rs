use criterion::{black_box, criterion_group, criterion_main, Criterion};
use spreadsheet_ods::{read_ods, OdsError};

fn read_orders() -> Result<(), OdsError> {
    let mut wb = read_ods("tests/orders.ods")?;
    Ok(())
}

fn criterion_benchmark(c: &mut Criterion) {
    c.bench_function("read orders", |b| b.iter(|| read_orders()));
}

criterion_group!(benches, criterion_benchmark);
criterion_main!(benches);
