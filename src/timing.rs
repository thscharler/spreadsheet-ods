use std::time::Instant;

pub fn timing<E>(name: &str, mut fun: impl FnMut() -> Result<(), E>) -> Result<(), E> {
    let now = Instant::now();
    fun()?;
    println!("{} {:?}", name, now.elapsed());
    Ok(())
}

pub fn timingn<E>(name: &str, mut fun: impl FnMut() -> ()) -> Result<(), E> {
    let now = Instant::now();
    fun();
    println!("{} {:?}", name, now.elapsed());
    Ok(())
}