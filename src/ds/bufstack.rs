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

    pub(crate) fn pop(&mut self) -> Vec<u8> {
        self.n += 1;
        self.buf.pop().unwrap_or_default()
    }

    pub(crate) fn push(&mut self, mut v: Vec<u8>) {
        self.n -= 1;
        v.clear();
        self.buf.push(v);
    }
}
