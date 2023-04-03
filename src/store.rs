use bytemuck::TransparentWrapper;
use std::collections::BTreeMap;

use slab::Slab;

#[derive(Debug, Clone, Copy, TransparentWrapper)]
#[repr(transparent)]
struct Position((u32, u32));

impl Position {
    #[inline]
    pub fn row(&self) -> u32 {
        self.0 .0
    }

    #[inline]
    pub fn col(&self) -> u32 {
        self.0 .1
    }
}

impl PartialEq for Position {
    #[inline]
    fn eq(&self, other: &Self) -> bool {
        self.row() == other.row() && self.col() == other.col()
    }
}

impl Eq for Position {}

impl PartialOrd for Position {
    #[inline]
    fn partial_cmp(&self, other: &Self) -> Option<std::cmp::Ordering> {
        Some(self.cmp(other))
    }
}

impl Ord for Position {
    fn cmp(&self, other: &Self) -> std::cmp::Ordering {
        if self.row() == other.row() {
            return self.col().cmp(&other.col());
        }
        self.row().cmp(&other.row())
    }
}

impl std::hash::Hash for Position {
    fn hash<H: std::hash::Hasher>(&self, state: &mut H) {
        self.0.hash(state);
    }
}

#[derive(Debug, Clone)]
pub struct Storage<T> {
    slab: Slab<T>,
    indexes: BTreeMap<Position, usize>,
}

impl<T> Default for Storage<T> {
    fn default() -> Self {
        Self {
            slab: Slab::new(),
            indexes: BTreeMap::new(),
        }
    }
}

impl<T> Storage<T> {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn insert(&mut self, pos: (u32, u32), val: T) {
        let index = self.slab.insert(val);
        self.indexes.insert(Position::wrap(pos), index);
    }

    pub fn get(&self, pos: &(u32, u32)) -> Option<&T> {
        self.indexes.get(Position::wrap_ref(pos)).map(|index| {
            self.slab
                .get(*index)
                .expect("Positions with an index must have a value in the slab")
        })
    }

    pub fn is_empty(&self) -> bool {
        self.indexes.is_empty()
    }

    pub fn iter(&self) -> Iter<'_, T> {
        let indexes = self.indexes.iter();
        Iter {
            store: self,
            indexes,
        }
    }
}

#[derive(Debug)]
pub struct Iter<'a, T> {
    store: &'a Storage<T>,
    indexes: std::collections::btree_map::Iter<'a, Position, usize>,
}

impl<'a, T> Iterator for Iter<'a, T> {
    type Item = (u32, u32, &'a T);

    fn next(&mut self) -> Option<Self::Item> {
        self.indexes.next().map(|(pos, index)| {
            let val = self
                .store
                .slab
                .get(*index)
                .expect("Positions with an index must have a value in the slab");

            (pos.row(), pos.col(), val)
        })
    }
}
