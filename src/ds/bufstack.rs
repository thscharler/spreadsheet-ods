pub(crate) struct BufStack {
    buf: Vec<Vec<u8>>,
}

impl BufStack {
    pub(crate) fn new() -> Self {
        BufStack { buf: Vec::new() }
    }

    pub(crate) fn get(&mut self) -> Vec<u8> {
        self.buf.pop().unwrap_or_else(Vec::new)
    }

    pub(crate) fn push(&mut self, v: Vec<u8>) {
        self.buf.push(v);
    }
}
