macro_rules! xlink_actuate {
    ($acc:ident) => {
        pub fn set_xlink_actuate(&mut self, actuate: XLinkActuate) {
            self.$acc.set_attr("xlink:actuate", actuate.to_string());
        }
    };
}

macro_rules! xlink_href {
    ($acc:ident) => {
        pub fn set_xlink_href<S: Into<String>>(&mut self, href: S) {
            self.$acc.set_attr("xlink:href", href.into());
        }
    };
}

macro_rules! xlink_show {
    ($acc:ident) => {
        pub fn set_xlink_show(&mut self, show: XLinkShow) {
            self.$acc.set_attr("xlink:show", show.to_string());
        }
    };
}

macro_rules! xlink_type {
    ($acc:ident) => {
        pub fn set_xlink_type(&mut self, ty: XLinkType) {
            self.$acc.set_attr("xlink:type", ty.to_string());
        }
    };
}
