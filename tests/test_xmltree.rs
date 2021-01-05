use spreadsheet_ods::xmltree::XmlTag;

#[test]
pub fn test_tree() {
    let _tag = XmlTag::new("table:shapes").tag(
        XmlTag::new("draw:frame")
            .attr("draw:z", "0")
            .attr("draw:name", "Bild 1")
            .attr("draw:styl:name", "gr1")
            .attr("draw:text-style-name", "P1")
            .attr("svg:width", "10.198cm")
            .attr("svg:height", "1.75cm")
            .attr("svg:x", "0cm")
            .attr("svg:y", "0cm")
            .tag(
                XmlTag::new("draw:image")
                    .attr(
                        "xlink:href",
                        "Pictures/10000000000011D7000003105281DD09B0E0B8D4.jpg",
                    )
                    .attr("xlink:type", "simple")
                    .attr("xlink:show", "embed")
                    .attr("xlink:actuate", "onLoad")
                    .attr("loext:mime-type", "image/jpeg")
                    .tag(XmlTag::new("text:p")),
            ),
    );
    // println!("{:?}", tag);
}
