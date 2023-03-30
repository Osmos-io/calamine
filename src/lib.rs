//! Rust Excel/OpenDocument reader
//!
//! # Status
//!
//! **calamine** is a pure Rust library to read Excel and OpenDocument Spreadsheet files.
//!
//! Read both cell values and vba project.
//!
//! # Examples
//! ```
//! use calamine::{Reader, open_workbook, Xlsx, DataType};
//!
//! // opens a new workbook
//! # let path = format!("{}/tests/issue3.xlsm", env!("CARGO_MANIFEST_DIR"));
//! let mut workbook: Xlsx<_> = open_workbook(path).expect("Cannot open file");
//!
//! // Read whole worksheet data and provide some statistics
//! if let Some(Ok(range)) = workbook.worksheet_range("Sheet1") {
//!     let total_cells = range.get_size().0 * range.get_size().1;
//!     let non_empty_cells: usize = range.used_cells().count();
//!     println!("Found {} cells in 'Sheet1', including {} non empty cells",
//!              total_cells, non_empty_cells);
//!     // alternatively, we can manually filter rows
//!     assert_eq!(non_empty_cells, range.rows()
//!         .flat_map(|r| r.into_iter().filter(|c| *c != &DataType::Empty)).count());
//! }
//!
//! // Check if the workbook has a vba project
//! if let Some(Ok(mut vba)) = workbook.vba_project() {
//!     let vba = vba.to_mut();
//!     let module1 = vba.get_module("Module 1").unwrap();
//!     println!("Module 1 code:");
//!     println!("{}", module1);
//!     for r in vba.get_references() {
//!         if r.is_missing() {
//!             println!("Reference {} is broken or not accessible", r.name);
//!         }
//!     }
//! }
//!
//! // You can also get defined names definition (string representation only)
//! for name in workbook.defined_names() {
//!     println!("name: {}, formula: {}", name.0, name.1);
//! }
//!
//! // Now get all formula!
//! let sheets = workbook.sheet_names().to_owned();
//! for s in sheets {
//!     println!("found {} formula in '{}'",
//!              workbook
//!                 .worksheet_formula(&s)
//!                 .expect("sheet not found")
//!                 .expect("error while getting formula")
//!                 .rows().flat_map(|r| r.into_iter().filter(|f| !f.is_empty()))
//!                 .count(),
//!              s);
//! }
//! ```
#![deny(missing_docs)]

#[macro_use]
mod utils;

mod auto;
mod cfb;
mod datatype;
mod ods;
mod xls;
mod xlsb;
mod xlsx;

mod de;
mod errors;
pub mod vba;

use serde::de::DeserializeOwned;
use std::borrow::Cow;
use std::cmp::{max, min};
use std::collections::HashMap;
use std::fmt;
use std::fs::File;
use std::io::{BufReader, Read, Seek};
use std::iter::{FusedIterator, Map};
use std::marker::PhantomData;
use std::ops::{Index, IndexMut};
use std::path::Path;

pub use crate::auto::{open_workbook_auto, open_workbook_auto_from_rs, Sheets};
pub use crate::datatype::DataType;
pub use crate::de::{DeError, RangeDeserializer, RangeDeserializerBuilder, ToCellDeserializer};
pub use crate::errors::Error;
pub use crate::ods::{Ods, OdsError};
pub use crate::xls::{Xls, XlsError, XlsOptions};
pub use crate::xlsb::{Xlsb, XlsbError};
pub use crate::xlsx::{OsmosXlsxConfig, Xlsx, XlsxError};

use crate::vba::VbaProject;

// https://msdn.microsoft.com/en-us/library/office/ff839168.aspx
/// An enum to represent all different errors that can appear as
/// a value in a worksheet cell
#[derive(Debug, Clone, PartialEq)]
pub enum CellErrorType {
    /// Division by 0 error
    Div0,
    /// Unavailable value error
    NA,
    /// Invalid name error
    Name,
    /// Null value error
    Null,
    /// Number error
    Num,
    /// Invalid cell reference error
    Ref,
    /// Value error
    Value,
    /// Getting data
    GettingData,
}

impl fmt::Display for CellErrorType {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> Result<(), fmt::Error> {
        match *self {
            CellErrorType::Div0 => write!(f, "#DIV/0!"),
            CellErrorType::NA => write!(f, "#N/A"),
            CellErrorType::Name => write!(f, "#NAME?"),
            CellErrorType::Null => write!(f, "#NULL!"),
            CellErrorType::Num => write!(f, "#NUM!"),
            CellErrorType::Ref => write!(f, "#REF!"),
            CellErrorType::Value => write!(f, "#VALUE!"),
            CellErrorType::GettingData => write!(f, "#DATA!"),
        }
    }
}

/// Common file metadata
///
/// Depending on file type, some extra information may be stored
/// in the Reader implementations
#[derive(Debug, Default)]
pub struct Metadata {
    sheets: Vec<String>,
    /// Map of sheet names/sheet path within zip archive
    names: Vec<(String, String)>,
}

// FIXME `Reader` must only be seek `Seek` for `Xls::xls`. Because of the present API this limits
// the kinds of readers (other) data in formats can be read from.
/// A trait to share spreadsheets reader functions across different `FileType`s
pub trait Reader<RS>: Sized
where
    RS: Read + Seek,
{
    /// Error specific to file type
    type Error: std::fmt::Debug + From<std::io::Error>;

    /// Creates a new instance.
    fn new(reader: RS) -> Result<Self, Self::Error>;
    /// Gets `VbaProject`
    fn vba_project(&mut self) -> Option<Result<Cow<'_, VbaProject>, Self::Error>>;
    /// Initialize
    fn metadata(&self) -> &Metadata;
    /// Read worksheet data in corresponding worksheet path
    fn worksheet_range(&mut self, name: &str) -> Option<Result<Range<DataType>, Self::Error>>;

    /// Fetch all worksheet data & paths
    fn worksheets(&mut self) -> Vec<(String, Range<DataType>)>;

    /// Read worksheet formula in corresponding worksheet path
    fn worksheet_formula(&mut self, _: &str) -> Option<Result<Range<String>, Self::Error>>;

    /// Get all sheet names of this workbook, in workbook order
    ///
    /// # Examples
    /// ```
    /// use calamine::{Xlsx, open_workbook, Reader};
    ///
    /// # let path = format!("{}/tests/issue3.xlsm", env!("CARGO_MANIFEST_DIR"));
    /// let mut workbook: Xlsx<_> = open_workbook(path).unwrap();
    /// println!("Sheets: {:#?}", workbook.sheet_names());
    /// ```
    fn sheet_names(&self) -> &[String] {
        &self.metadata().sheets
    }

    /// Get all defined names (Ranges names etc)
    fn defined_names(&self) -> &[(String, String)] {
        &self.metadata().names
    }

    /// Get the nth worksheet. Shortcut for getting the nth
    /// sheet_name, then the corresponding worksheet.
    fn worksheet_range_at(&mut self, n: usize) -> Option<Result<Range<DataType>, Self::Error>> {
        let name = self.sheet_names().get(n)?.to_string();
        self.worksheet_range(&name)
    }
}

/// Convenient function to open a file with a BufReader<File>
pub fn open_workbook<R, P>(path: P) -> Result<R, R::Error>
where
    R: Reader<BufReader<File>>,
    P: AsRef<Path>,
{
    let file = BufReader::new(File::open(path)?);
    R::new(file)
}

/// Convenient function to open a file with a BufReader<File>
pub fn open_workbook_from_rs<R, RS>(rs: RS) -> Result<R, R::Error>
where
    RS: Read + Seek,
    R: Reader<RS>,
{
    R::new(rs)
}

/// A trait to constrain cells
pub trait CellType: Default + Clone + PartialEq {}

impl CellType for DataType {}
impl CellType for String {}
impl CellType for usize {} // for tests

/// A struct to hold cell position and value
#[derive(Debug, Clone)]
pub struct Cell<T: CellType> {
    /// Position for the cell (row, column)
    pos: (u32, u32),
    /// Value for the cell
    val: T,
}

impl<T: CellType> Cell<T> {
    /// Creates a new `Cell`
    pub fn new(position: (u32, u32), value: T) -> Cell<T> {
        Cell {
            pos: position,
            val: value,
        }
    }

    /// Gets `Cell` position
    pub fn get_position(&self) -> (u32, u32) {
        self.pos
    }

    /// Gets `Cell` value
    pub fn get_value(&self) -> &T {
        &self.val
    }
}

/// A struct which represents a squared selection of cells
#[derive(Debug, Default, Clone)]
pub struct Range<T> {
    start: (u32, u32),
    end: (u32, u32),
    inner: HashMap<(u32, u32), T>,
    empty_value: T,
}

impl<T> Range<T> {
    /// Get top left cell position (row, column)
    #[inline]
    pub fn start(&self) -> Option<(u32, u32)> {
        if self.is_empty() {
            None
        } else {
            Some(self.start)
        }
    }

    /// Get bottom right cell position (row, column)
    #[inline]
    pub fn end(&self) -> Option<(u32, u32)> {
        if self.is_empty() {
            None
        } else {
            Some(self.end)
        }
    }

    /// Get column width
    #[inline]
    pub fn width(&self) -> usize {
        if self.is_empty() {
            0
        } else {
            (self.end.1 - self.start.1 + 1) as usize
        }
    }

    /// Get column height
    #[inline]
    pub fn height(&self) -> usize {
        if self.is_empty() {
            0
        } else {
            (self.end.0 - self.start.0 + 1) as usize
        }
    }

    /// Get size in (height, width) format
    #[inline]
    pub fn get_size(&self) -> (usize, usize) {
        (self.height(), self.width())
    }

    /// Is range empty
    #[inline]
    pub fn is_empty(&self) -> bool {
        self.inner.is_empty()
    }

    /// Set inner value from absolute position
    ///
    /// # Remarks
    ///
    /// Will try to resize inner structure if the value is out of bounds.
    /// For relative positions, use Index trait
    ///
    /// Try to avoid this method as much as possible and prefer initializing
    /// the `Range` with `from_sparse` constructor.
    ///
    /// # Panics
    ///
    /// If absolute_position > Cell start
    ///
    /// # Examples
    /// ```
    /// use calamine::{Range, DataType};
    ///
    /// let mut range = Range::new((0, 0), (5, 2));
    /// assert_eq!(range.get_value((2, 1)), Some(&DataType::Empty));
    /// range.set_value((2, 1), DataType::Float(1.0));
    /// assert_eq!(range.get_value((2, 1)), Some(&DataType::Float(1.0)));
    /// ```
    pub fn set_value(&mut self, absolute_position: (u32, u32), value: T) {
        assert!(
            self.start.0 <= absolute_position.0 && self.start.1 <= absolute_position.1,
            "absolute_position out of bounds"
        );
        // maybe update the start and end values
        self.start.0 = self.start.0.min(absolute_position.0);
        self.start.1 = self.start.1.min(absolute_position.1);

        self.end.0 = self.end.0.max(absolute_position.0);
        self.end.1 = self.end.1.max(absolute_position.1);

        // add the value into matrix
        self.inner.insert(absolute_position, value);
    }

    /// Get cell value from **absolute position**.
    ///
    /// If the `absolute_position` is out of range, returns `None`, else returns the cell value.
    /// The coordinate format is (row, column).
    ///
    /// # Warnings
    ///
    /// For relative positions, use Index trait
    ///
    /// # Remarks
    ///
    /// Absolute position is in *sheet* referential while relative position is in *range* referential.
    ///
    /// For instance if we consider range *C2:H38*:
    /// - `(0, 0)` absolute is "A1" and thus this function returns `None`
    /// - `(0, 0)` relative is "C2" and is returned by the `Index` trait (i.e `my_range[(0, 0)]`)
    ///
    /// # Examples
    /// ```
    /// use calamine::{Range, DataType};
    ///
    /// let range: Range<usize> = Range::new((1, 0), (5, 2));
    /// assert_eq!(range.get_value((0, 0)), None);
    /// assert_eq!(range[(0, 0)], 0);
    /// ```
    #[inline]
    pub fn get_value(&self, absolute_position: (u32, u32)) -> Option<&T> {
        if self.start.0 > absolute_position.0
            || self.start.1 > absolute_position.1
            || self.end.0 < absolute_position.0
            || self.end.1 < absolute_position.1
        {
            return None;
        }

        Some(
            self.inner
                .get(&absolute_position)
                .unwrap_or(&self.empty_value),
        )
    }

    /// Get cell value from **relative position**.
    pub fn get(&self, relative_position: (usize, usize)) -> Option<&T> {
        let absolute_position = (
            relative_position.0 as u32 + self.start.0,
            relative_position.1 as u32 + self.start.1,
        );

        self.get_value(absolute_position)
    }

    /// Get an iterator over inner rows
    ///
    /// # Examples
    /// ```
    /// use calamine::{Range, DataType};
    ///
    /// let range: Range<DataType> = Range::new((0, 0), (5, 2));
    /// // with rows item row: &[DataType]
    /// assert_eq!(range.rows().map(|r| r.len()).sum::<usize>(), 18);
    /// ```
    pub fn rows(&self) -> Rows<'_, T> {
        // if self.inner.is_empty() {
        //     Rows { inner: None }
        // } else {
        //     let width = self.width();
        //     Rows {
        //         inner: Some(self.inner.chunks(width)),
        //     }
        // }
        Rows::new(self)
    }

    /// Get an iterator over used cells only
    pub fn used_cells<'a>(&'a self) -> UsedCells<'a, T> {
        fn to_triplet<'a, T>(((row, col), val): (&'a (u32, u32), &'a T)) -> (usize, usize, &'a T) {
            (*row as usize, *col as usize, val)
        }

        let inner = self
            .inner
            .iter()
            .map(to_triplet as fn((&'a (u32, u32), &'a T)) -> (usize, usize, &'a T));

        UsedCells { inner }
    }

    /// Get an iterator over all cells in this range
    pub fn cells(&self) -> Cells<'_, T> {
        Cells::new(self)
    }

    /// Build a `RangeDeserializer` from this configuration.
    ///
    /// # Example
    ///
    /// ```
    /// # use calamine::{Reader, Error, open_workbook, Xlsx, RangeDeserializerBuilder};
    /// fn main() -> Result<(), Error> {
    ///     let path = format!("{}/tests/temperature.xlsx", env!("CARGO_MANIFEST_DIR"));
    ///     let mut workbook: Xlsx<_> = open_workbook(path)?;
    ///     let mut sheet = workbook.worksheet_range("Sheet1")
    ///         .ok_or(Error::Msg("Cannot find 'Sheet1'"))??;
    ///     let mut iter = sheet.deserialize()?;
    ///
    ///     if let Some(result) = iter.next() {
    ///         let (label, value): (String, f64) = result?;
    ///         assert_eq!(label, "celsius");
    ///         assert_eq!(value, 22.2222);
    ///
    ///         Ok(())
    ///     } else {
    ///         return Err(From::from("expected at least one record but got none"));
    ///     }
    /// }
    /// ```
    pub fn deserialize<'a, D>(&'a self) -> Result<RangeDeserializer<'a, T, D>, DeError>
    where
        T: ToCellDeserializer<'a>,
        D: DeserializeOwned,
    {
        RangeDeserializerBuilder::new().from_range(self)
    }
}

impl<T: Default> Range<T> {
    /// Creates a new non-empty `Range`
    ///
    /// When possible, prefer the more efficient `Range::from_sparse`
    ///
    /// # Panics
    ///
    /// Panics if start.0 > end.0 or start.1 > end.1
    #[inline]
    pub fn new(start: (u32, u32), end: (u32, u32)) -> Range<T> {
        assert!(start <= end, "invalid range bounds");
        Range {
            start,
            end,
            inner: HashMap::default(),
            empty_value: T::default(),
        }
    }

    /// Creates a new empty range
    #[inline]
    pub fn empty() -> Range<T> {
        Range {
            start: (0, 0),
            end: (0, 0),
            inner: HashMap::default(),
            empty_value: T::default(),
        }
    }

    /// Create a range from a list of triplets in the form `(row, col, T)`.
    pub fn from_triplets<I: Iterator<Item = (u32, u32, T)>>(triplets: I) -> Self {
        let (start, end, inner) = triplets.fold(
            ((u32::MAX, u32::MIN), (0u32, 0u32), HashMap::new()),
            |(mut start, mut end, mut inner), (row, col, val)| {
                start.0 = min(start.0, row);
                start.1 = min(start.1, col);

                end.0 = max(end.0, row);
                end.1 = max(end.1, col);

                inner.insert((row, col), val);

                (start, end, inner)
            },
        );

        Self {
            start,
            end,
            inner,
            empty_value: T::default(),
        }
    }
}

impl<T: Clone> Range<T> {
    /// Build a new `Range` out of this range
    ///
    /// # Remarks
    ///
    /// Cells within this range will be cloned, cells out of it will be set to Empty
    ///
    /// # Example
    ///
    /// ```
    /// # use calamine::{Range, DataType};
    /// let mut a = Range::new((1, 1), (3, 3));
    /// a.set_value((1, 1), DataType::Bool(true));
    /// a.set_value((2, 2), DataType::Bool(true));
    ///
    /// let b = a.range((2, 2), (5, 5));
    /// assert_eq!(b.get_value((2, 2)), Some(&DataType::Bool(true)));
    /// assert_eq!(b.get_value((3, 3)), Some(&DataType::Empty));
    ///
    /// let c = a.range((0, 0), (2, 2));
    /// assert_eq!(c.get_value((0, 0)), Some(&DataType::Empty));
    /// assert_eq!(c.get_value((1, 1)), Some(&DataType::Bool(true)));
    /// assert_eq!(c.get_value((2, 2)), Some(&DataType::Bool(true)));
    /// ```
    pub fn range(&self, start: (u32, u32), end: (u32, u32)) -> Range<T> {
        let subview = self
            .inner
            .clone()
            .into_iter()
            .filter(|(pos, _value)| pos >= &start && pos <= &end)
            .collect::<HashMap<(u32, u32), T>>();

        Self {
            start,
            end,
            inner: subview,
            empty_value: self.empty_value.clone(),
        }
    }
}

impl<T: CellType> Range<T> {
    /// Creates a `Range` from a coo sparse vector of `Cell`s.
    ///
    /// Coordinate list (COO) is the natural way cells are stored
    /// Inner size is defined only by non empty.
    ///
    /// cells: `Vec` of non empty `Cell`s, sorted by row
    ///
    /// # Panics
    ///
    /// panics when a `Cell` row is lower than the first `Cell` row or
    /// bigger than the last `Cell` row.
    pub fn from_sparse(cells: Vec<Cell<T>>) -> Range<T> {
        if cells.is_empty() {
            Range::empty()
        } else {
            let capacity = cells.len();

            let (start, end, inner) = cells.into_iter().fold(
                (
                    (u32::MAX, u32::MAX),
                    (0u32, 0u32),
                    HashMap::with_capacity(capacity),
                ),
                |(mut start, mut end, mut inner), cell| {
                    if start.0 > cell.pos.0 {
                        start.0 = cell.pos.0;
                    }
                    if start.1 > cell.pos.1 {
                        start.1 = cell.pos.1;
                    }
                    if end.0 < cell.pos.0 {
                        end.0 = cell.pos.0;
                    }
                    if end.1 < cell.pos.1 {
                        end.1 = cell.pos.1;
                    }
                    inner.insert(cell.pos, cell.val);
                    (start, end, inner)
                },
            );

            Range {
                start,
                end,
                inner,
                empty_value: T::default(),
            }
        }
    }
}

impl<T: CellType> Index<(usize, usize)> for Range<T> {
    type Output = T;
    fn index(&self, (row, col): (usize, usize)) -> &T {
        self.get_value((row as u32, col as u32))
            .unwrap_or(&self.empty_value)
    }
}

/// Create an iterator over the coordinates in the start and end range.
#[derive(Debug)]
pub struct IndexIter {
    start: (u32, u32),
    end: (u32, u32),
    row: u32,
    col: u32,
}

impl IndexIter {
    /// Create a new iterator over the coordinates in the start and end range.
    pub fn new(start: (u32, u32), end: (u32, u32)) -> Self {
        let row = start.0;
        let col = start.1;

        Self {
            start,
            end,
            row,
            col,
        }
    }

    /// Get the number of columns in the iter. This is the number of values per row.
    pub fn num_columns(&self) -> usize {
        (self.end.1 - self.start.1) as usize
    }

    /// Get the number of rows in the iterator
    pub fn num_rows(&self) -> usize {
        (self.end.0 - self.start.0) as usize
    }
}

impl Iterator for IndexIter {
    type Item = (u32, u32);

    fn next(&mut self) -> Option<Self::Item> {
        if self.row > self.end.0 {
            return None;
        }

        let pos = (self.row, self.col);

        self.col += 1;

        if self.col > self.end.1 {
            self.col = self.start.1;
            self.row += 1;
        }

        Some(pos)
    }
}

impl FusedIterator for IndexIter {}

/// Iterator over cells within some container with a known size.
pub struct SizedCellIter<'a, T> {
    len: usize,
    range: &'a Range<T>,
    indexes: Box<dyn Iterator<Item = (u32, u32)>>,
}

impl<'a, T> SizedCellIter<'a, T> {
    /// Create an iterator over a set of cells in a range
    pub fn new<'b: 'a>(
        range: &'a Range<T>,
        indexes: Box<dyn Iterator<Item = (u32, u32)>>,
        len: usize,
    ) -> Self {
        Self {
            len,
            indexes,
            range,
        }
    }
}

impl<'a, T> Iterator for SizedCellIter<'a, T> {
    type Item = &'a T;

    fn next(&mut self) -> Option<Self::Item> {
        if let Some(pos) = self.indexes.next() {
            self.range.get_value(pos)
        } else {
            None
        }
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        (self.len, Some(self.len))
    }
}

impl<'a, T> ExactSizeIterator for SizedCellIter<'a, T> {}

/// Iterator over the rows of a range.
pub struct Rows<'a, T> {
    range: &'a Range<T>,
    current_row: u32,
    bounds: ((u32, u32), (u32, u32)),
}

impl<'a, T> Rows<'a, T> {
    /// Create a new iterator over the rows of a range.
    pub fn new(range: &'a Range<T>) -> Self {
        let start = range.start;
        let end = range.end;

        Self {
            range,
            current_row: start.0,
            bounds: (start, end),
        }
    }
}

impl<'a, T> Iterator for Rows<'a, T> {
    type Item = SizedCellIter<'a, T>;

    fn next(&mut self) -> Option<Self::Item> {
        let current_row = self.current_row;
        self.current_row = current_row + 1;

        let (start, end) = self.bounds;

        if current_row > end.0 {
            return None;
        }

        Some(SizedCellIter::new(
            self.range,
            Box::new((start.1..=end.1).map(move |col| (current_row, col))),
            (end.1 - start.1) as usize + 1,
        ))
    }
}

/// A struct to iterate over all cells
#[derive(Debug)]
pub struct Cells<'a, T> {
    range: &'a Range<T>,
    iter: IndexIter,
}

impl<'a, T> Cells<'a, T> {
    /// Create a new iterator over all cells in the range.
    pub fn new(range: &'a Range<T>) -> Self {
        let iter = IndexIter::new(range.start, range.end);

        Self { range, iter }
    }
}

impl<'a, T: 'a + CellType> Iterator for Cells<'a, T> {
    type Item = (usize, usize, &'a T);

    fn next(&mut self) -> Option<Self::Item> {
        self.iter.next().map(|(row, col)| {
            let val = self
                .range
                .get_value((row, col))
                .expect("IndexIter produced an invalid (row, col) combination");
            (row as usize, col as usize, val)
        })
    }
}

type TripletFn<'a, T> = fn((&'a (u32, u32), &'a T)) -> (usize, usize, &'a T);
type UsedCellIter<'a, T> =
    Map<std::collections::hash_map::Iter<'a, (u32, u32), T>, TripletFn<'a, T>>;

/// A struct to iterate over used cells
#[derive(Debug)]
pub struct UsedCells<'a, T> {
    inner: UsedCellIter<'a, T>,
}

impl<'a, T> Iterator for UsedCells<'a, T>
where
    T: 'a + CellType,
{
    type Item = (usize, usize, &'a T);

    fn next(&mut self) -> Option<Self::Item> {
        self.inner.next()
    }
}

/// Struct with the key elements of a table
pub struct Table<T> {
    pub(crate) name: String,
    pub(crate) sheet_name: String,
    pub(crate) columns: Vec<String>,
    pub(crate) data: Range<T>,
}

impl<T> Table<T> {
    /// Get the name of the table
    pub fn name(&self) -> &str {
        &self.name
    }
    /// Get the name of the sheet that table exists within
    pub fn sheet_name(&self) -> &str {
        &self.sheet_name
    }
    /// Get the names of the columns in the order they occur
    pub fn columns(&self) -> &[String] {
        &self.columns
    }
    /// Get a range representing the data from the table (excludes column headers)
    pub fn data(&self) -> &Range<T> {
        &self.data
    }
}

#[cfg(test)]
mod test {
    use crate::IndexIter;

    #[test]
    fn test_index_iter() {
        let start = (0, 0);
        let end = (2, 2);

        let mut iter = IndexIter::new(start, end);

        assert_eq!(iter.next(), Some((0, 0)));
        assert_eq!(iter.next(), Some((0, 1)));
        assert_eq!(iter.next(), Some((0, 2)));

        assert_eq!(iter.next(), Some((1, 0)));
        assert_eq!(iter.next(), Some((1, 1)));
        assert_eq!(iter.next(), Some((1, 2)));

        assert_eq!(iter.next(), Some((2, 0)));
        assert_eq!(iter.next(), Some((2, 1)));
        assert_eq!(iter.next(), Some((2, 2)));

        assert_eq!(iter.next(), None);
    }
}
