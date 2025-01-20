/// deg angles. 360°
#[macro_export]
macro_rules! deg {
    ($l:expr) => {
        $crate::Angle::Deg($l.into()).into()
    };
}

/// grad angles. 400°
#[macro_export]
macro_rules! grad {
    ($l:expr) => {
        $crate::Angle::Grad($l.into()).into()
    };
}

/// radians angle.
#[macro_export]
macro_rules! rad {
    ($l:expr) => {
        $crate::Angle::Rad($l.into()).into()
    };
}

/// Centimeters.
#[macro_export]
macro_rules! cm {
    ($l:expr) => {
        $crate::Length::Cm($l.into()).into()
    };
}

/// Millimeters.
#[macro_export]
macro_rules! mm {
    ($l:expr) => {
        $crate::Length::Mm($l.into()).into()
    };
}

/// Inches.
#[macro_export]
macro_rules! inch {
    ($l:expr) => {
        $crate::Length::In($l.into()).into()
    };
}

/// Point. 1/72"
#[macro_export]
macro_rules! pt {
    ($l:expr) => {
        $crate::Length::Pt($l.into()).into()
    };
}

/// Pica. 12/72"
#[macro_export]
macro_rules! pc {
    ($l:expr) => {
        $crate::Length::Pc($l into()).into()
    };
}

/// Length depending on font size.
#[macro_export]
macro_rules! em {
    ($l:expr) => {
        $crate::Length::Em($l into()).into()
    };
}
