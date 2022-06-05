pub(crate) struct BufStack {
    n: i32,
    buf: Vec<Vec<u8>>,
}

impl BufStack {
    pub(crate) fn new() -> Self {
        BufStack {
            n: 0,
            buf: Vec::new(),
        }
    }

    pub(crate) fn get_buf(&mut self) -> Vec<u8> {
        self.n += 1;
        self.buf.pop().unwrap_or_default()
    }

    pub(crate) fn push(&mut self, v: Vec<u8>) {
        self.n -= 1;
        self.buf.push(v);
    }
}
