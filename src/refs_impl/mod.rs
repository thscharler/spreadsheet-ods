use crate::refs_impl::error::{OFCode, ParseOFError};
use nom::Offset;
use nom_locate::LocatedSpan;
use std::slice;

pub(crate) mod ast;
pub(crate) mod parser;

// clones from openformula
#[allow(unreachable_pub)]
#[allow(unused)]
pub(crate) mod conv;
#[allow(unused)]
#[allow(unreachable_pub)]
pub(crate) mod error;
#[allow(unused)]
#[allow(unreachable_pub)]
pub(crate) mod format;
#[allow(unused)]
#[allow(unreachable_pub)]
pub(crate) mod tokens;

/// Input type.
pub(crate) type Span<'a> = LocatedSpan<&'a str>;

/// Result type.
pub(crate) type ParseResult<'s, O> = Result<(Span<'s>, O), ParseOFError<'s>>;

/// Returns a new Span that reaches from the beginning of span0 to the end of span1.
///
/// If any of the following conditions are violated, the result is Undefined Behavior:
/// * Both the starting and other pointer must be either in bounds or one byte past the end of the same allocated object.
///      Should be guaranteed if both were obtained from on ast run.
/// * Both pointers must be derived from a pointer to the same object.
///      Should be guaranteed if both were obtained from on ast run.
/// * The distance between the pointers, in bytes, cannot overflow an isize.
/// * The distance being in bounds cannot rely on “wrapping around” the address space.
pub(crate) unsafe fn span_union<'a>(span0: Span<'a>, span1: Span<'a>) -> Span<'a> {
    let ptr = span0.as_ptr();
    // offset to the start of span1 and add the length of span1.
    let size = span0.offset(&span1) + span1.len();

    unsafe {
        // The size should be within the original allocation, if both spans are from
        // the same ast run. We must ensure that the ast run doesn't generate
        // Spans out of nothing that end in the ast.
        let slice = slice::from_raw_parts(ptr, size);
        // This is all from a str originally and we never got down to bytes.
        let str = std::str::from_utf8_unchecked(slice);

        // As span0 was ok the offset used here is ok too.
        Span::new_from_raw_offset(span0.location_offset(), span0.location_line(), str, ())
    }
}

/// Fails if the string was not fully parsed.
pub(crate) fn check_eof<'s>(i: Span<'s>, code: OFCode) -> Result<(), ParseOFError<'s>> {
    if (*i).is_empty() {
        Ok(())
    } else {
        Err(ParseOFError::new(code, i))
    }
}

/// Change the error code.
pub(crate) fn map_err<'s, O>(mut err: ParseOFError<'s>, code: OFCode) -> ParseResult<'s, O> {
    // Translates the code with some exceptions.
    if err.code != OFCode::OFCNomError
        && err.code != OFCode::OFCNomFailure
        && err.code != OFCode::OFCUnexpected
        && err.code != OFCode::OFCParseIncomplete
    {
        err.code = code;
    }

    Err(err)
}

/// Panic in the parser.
#[track_caller]
pub(crate) fn panic_parse<'s>(e: ParseOFError<'s>) -> ! {
    unreachable!("{}", e)
}

/// MOD: Here it just calls map_err to change the error-code in a convenient way.
///
/// --Helps with keeping tracks in the parsers.
///
/// This can be squeezed between the call to another parser and the ?-operator.
///
/// Makes sure the tracer can keep track of the complete parse call tree.--
pub(crate) trait TrackParseResult<'s, 't, O> {
    /// Translates the error code and adds the standard expect value.
    /// Then tracks the error and marks the current function as finished.
    fn trace(self, code: OFCode) -> Self;
}

impl<'s, 't, O> TrackParseResult<'s, 't, O> for ParseResult<'s, O> {
    fn trace(self, code: OFCode) -> Self {
        match self {
            Ok(_) => self,
            Err(e) => map_err(e, code),
        }
    }
}
