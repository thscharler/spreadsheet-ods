use spreadsheet_ods::xmltree::XmlTag;

#[test]
pub fn test_tree() {
    let tag = XmlTag::new("table:shapes")
        .con_tag(XmlTag::new("draw:frame")
            .con_attr("draw:z", "0")
            .con_attr("draw:name", "Bild 1")
            .con_attr("draw:styl:name", "gr1")
            .con_attr("draw:text-style-name", "P1")
            .con_attr("svg:width", "10.198cm")
            .con_attr("svg:height", "1.75cm")
            .con_attr("svg:x", "0cm")
            .con_attr("svg:y", "0cm")
            .con_tag(XmlTag::new("draw:image")
                .con_attr("xlink:href", "Pictures/10000000000011D7000003105281DD09B0E0B8D4.jpg")
                .con_attr("xlink:type", "simple")
                .con_attr("xlink:show", "embed")
                .con_attr("xlink:actuate", "onLoad")
                .con_attr("loext:mime-type", "image/jpeg")
                .con_tag(XmlTag::new("text:p"))
            )
        );
    println!("{:?}", tag);
}