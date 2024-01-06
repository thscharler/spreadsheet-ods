use smol_str::SmolStr;

pub(crate) mod detach;

pub(crate) fn size_of_smolstr(str: &SmolStr) -> usize {
    if str.is_heap_allocated() {
        str.len()
    } else {
        0
    }
}
