use spreadsheet_ods::text;
use spreadsheet_ods::text::{AuthorName, TextH, TextTag};
use spreadsheet_ods::xmltree::XmlTag;

#[test]
pub fn test_tree() {
    let tag = XmlTag::new("table:shapes").tag(
        XmlTag::new("draw:frame")
            .attr_slice(&[
                ("draw:z", "0".into()),
                ("draw:name", "Bild 1".into()),
                ("draw:styl:name", "gr1".into()),
                ("draw:text-style-name", "P1".into()),
                ("svg:width", "10.198cm".into()),
                ("svg:height", "1.75cm".into()),
                ("svg:x", "0cm".into()),
                ("svg:y", "0cm".into()),
            ])
            .tag(
                XmlTag::new("draw:image")
                    .attr_slice(&[
                        (
                            "xlink:href",
                            "Pictures/10000000000011D7000003105281DD09B0E0B8D4.jpg".into(),
                        ),
                        ("xlink:type", "simple".into()),
                        ("xlink:show", "embed".into()),
                        ("xlink:actuate", "onLoad".into()),
                        ("loext:mime-type", "image/jpeg".into()),
                    ])
                    .tag(XmlTag::new("text:p")),
            ),
    );
    println!("{}", tag);
}

#[test]
pub fn test_text() {
    let txt = TextTag::new("text:p")
        .tag(AuthorName::new())
        .tag(TextH::new().style_name(&"flfl".into()).text("heyder"));
    println!("{:?}", txt);
    println!("{}", txt);
}
