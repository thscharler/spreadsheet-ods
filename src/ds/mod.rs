use loupe::{MemoryUsage, MemoryUsageTracker};
use std::mem;

pub(crate) mod detach;

// Copied from <https://github.com/Aeledfyr/deepsize>

// A btree node has 2*B - 1 (K,V) pairs and (usize, u16, u16)
// overhead, and an internal btree node additionally has 2*B
// `usize` overhead.
// A node can contain between B - 1 and 2*B - 1 elements, so
// we assume it has the midpoint 3/2*B - 1.

// Constants from rust's source:
// https://doc.rust-lang.org/src/alloc/collections/btree/node.rs.html#43-45

const BTREE_B: usize = 6;
const BTREE_MIN: usize = 2 * BTREE_B - 1;
const BTREE_MAX: usize = BTREE_B - 1;

pub(crate) fn size_of_btreemap<K: MemoryUsage, V: MemoryUsage>(
    map: &std::collections::BTreeMap<K, V>,
    tracker: &mut dyn MemoryUsageTracker,
) -> usize {
    let element_size = map.iter().fold(0, |sum, (k, v)| {
        sum + k.size_of_val(tracker) + v.size_of_val(tracker)
    });
    let overhead = mem::size_of::<(usize, u16, u16, [(K, V); BTREE_MAX], [usize; BTREE_B])>();

    element_size + map.len() * overhead * 2 / (BTREE_MAX + BTREE_MIN)
}
