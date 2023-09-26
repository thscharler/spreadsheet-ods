//! Row and column groups.

/// Row group
#[derive(Debug, PartialEq, Clone)]
pub struct RowGroup {
    from: u32,
    to: u32,
    display: bool,
}

impl RowGroup {
    ///
    pub fn new(from: u32, to: u32, display: bool) -> Self {
        Self { from, to, display }
    }

    /// Inclusive start row.
    pub fn from(&self) -> u32 {
        self.from
    }

    /// Inclusive start row.
    pub fn set_from(&mut self, from: u32) {
        self.from = from;
    }

    /// Inclusive end row.
    pub fn to(&self) -> u32 {
        self.to
    }

    /// Inclusive end row.
    pub fn set_to(&mut self, to: u32) {
        self.to = to
    }

    /// Contains the other group.
    pub fn contains(&self, other: &RowGroup) -> bool {
        self.from <= other.from && self.to >= other.to
    }

    /// The two ranges are disjunct.
    pub fn disjunct(&self, other: &RowGroup) -> bool {
        self.from < other.from && self.to < other.from || self.from > other.to && self.to > other.to
    }

    /// Rowgroup is expanded?
    pub fn display(&self) -> bool {
        self.display
    }

    /// Set the expanded state for the row group.
    ///
    /// Note: Changing this does not change the visibility of the rows.
    /// Use Sheet::collapse_row_group() to make all necessary changes.
    pub fn set_display(&mut self, display: bool) {
        self.display = display;
    }
}

/// Column groups.
#[derive(Debug, PartialEq, Clone)]
pub struct ColGroup {
    from: u32,
    to: u32,
    display: bool,
}

impl ColGroup {
    ///
    pub fn new(from: u32, to: u32, display: bool) -> Self {
        Self { from, to, display }
    }

    /// Inclusive start row.
    pub fn from(&self) -> u32 {
        self.from
    }

    /// Inclusive start row.
    pub fn set_from(&mut self, from: u32) {
        self.from = from;
    }

    /// Inclusive end row.
    pub fn to(&self) -> u32 {
        self.to
    }

    /// Inclusive end row.
    pub fn set_to(&mut self, to: u32) {
        self.to = to
    }

    /// Contains the other group.
    pub fn contains(&self, other: &ColGroup) -> bool {
        self.from <= other.from && self.to >= other.to
    }

    /// The two ranges are disjunct.
    pub fn disjunct(&self, other: &ColGroup) -> bool {
        self.from < other.from && self.to < other.from || self.from > other.to && self.to > other.to
    }

    /// Col group is expanded?
    pub fn display(&self) -> bool {
        self.display
    }

    /// Change the expanded state for the col group.
    ///
    /// Note: Changing this does not change the visibility of the columns.
    /// Use Sheet::collapse_col_group() to make all necessary changes.
    pub fn set_display(&mut self, display: bool) {
        self.display = display;
    }
}
