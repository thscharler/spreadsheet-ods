use spreadsheet_ods::xmltree::XmlTag;

#[test]
pub fn test_tree() {
    let tag = XmlTag::tag("table:shapes")
        .tag_con(XmlTag::tag("draw:frame")
            .attr_con("draw:z", "0")
            .attr_con("draw:name", "Bild 1")
            .attr_con("draw:styl:name", "gr1")
            .attr_con("draw:text-style-name", "P1")
            .attr_con("svg:width", "10.198cm")
            .attr_con("svg:height", "1.75cm")
            .attr_con("svg:x", "0cm")
            .attr_con("svg:y", "0cm")
            .tag_con(XmlTag::tag("draw:image")
                .attr_con("xlink:href", "Pictures/10000000000011D7000003105281DD09B0E0B8D4.jpg")
                .attr_con("xlink:type", "simple")
                .attr_con("xlink:show", "embed")
                .attr_con("xlink:actuate", "onLoad")
                .attr_con("loext:mime-type", "image/jpeg")
                .tag_con(XmlTag::tag("text:p"))
            )
        );
    println!("{:?}", tag);
}