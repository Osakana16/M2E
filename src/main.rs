use std::{borrow::Borrow, io::Cursor, iter::Filter};

use image::{codecs::png::{self, FilterType, PngDecoder}, DynamicImage};
use xlsxwriter::{worksheet::{ImageOptions, WorksheetCol, WorksheetRow}, Format, Workbook, Worksheet};


fn main() {
    let args: Vec<String> = std::env::args().collect();
    compile_excel(&args[1], &args[2]);
}

fn compile_excel(markdown_file_name: &str, xlsx_file_name: &str) {
    let ast = markdown::to_mdast(&std::fs::read_to_string(markdown_file_name).unwrap(), &markdown::ParseOptions::default());
    println!("{:?}", ast);

    let workbook = Workbook::new(xlsx_file_name);
    match workbook {
        Ok(w) =>  
            match ast { 
                Ok(a) => make_excel(w, a),
                Err(_) => {}
            },           
        Err(_) => {}
    };
}

// - Markdown -

fn expand_ast(node: markdown::mdast::Node, mut sheet: &mut Worksheet, x: WorksheetCol, y: WorksheetRow, format: Option<Format>) {
    match node.clone() {
        markdown::mdast::Node::Root(r) => {
            for i in 0..r.children.len() {
                expand_ast_children(&r.children, i, sheet, 0, (3 + i).try_into().unwrap(), None)
            }
        },
        markdown::mdast::Node::Paragraph(p) => {
            for i in 0..p.children.len() {
                expand_ast_children(&p.children, i, sheet, (i * 2) as WorksheetCol, y, format.clone());
            }
        },
        markdown::mdast::Node::Heading(h) => {
            match h.depth {
                1 => expand_ast_children(&h.children, 0, sheet, 0, y, Some(create_h1_format())),
                2 => expand_ast_children(&h.children, 0, sheet, 0, y, Some(create_h2_format())),
                _ => expand_ast_children(&h.children, 0, sheet, 0, y, Some(create_h3_format()))
            };
        },
        markdown::mdast::Node::Text(t) => write_string(&mut sheet, &t.value, x, y,  format),
        markdown::mdast::Node::Code(c) => write_string(&mut sheet, &c.value, 0, y, Some(create_code_format())),
        markdown::mdast::Node::Emphasis(e) => expand_ast_children(&e.children, 0, sheet, x, y, Some(create_bold_format())),
        markdown::mdast::Node::InlineCode(ic) => write_string(&mut sheet, &ic.value, x, y, Some(create_inline_format())),
        markdown::mdast::Node::BlockQuote(bc) => expand_ast_children(&bc.children, 0, sheet, 0, y, Some(create_quotation_format())),
        markdown::mdast::Node::Image(i) => insert_image(&mut sheet, &i.url, y),
        _ => ()
    };
}

fn expand_ast_children(children: &Vec<markdown::mdast::Node>, i: usize, mut sheet: &mut Worksheet, x: WorksheetCol, y: WorksheetRow, format: Option<Format>) {
    expand_ast(children[i].clone(), sheet, x, y, format);
}

// - Excel -

fn make_excel(workbook: Workbook, ast: markdown::mdast::Node) {
    let sheet = workbook.add_worksheet(Some("sheet1"));
    match sheet {
        Ok(mut s) => {
            init_sheet(&mut s);
            let max_row = ast.children().unwrap().len() + 3;
            for i in 3..max_row {
                s.set_row(i.try_into().unwrap(), 28.0, None);
            }
            expand_ast(ast, &mut s, 0, 0, None);            
        },
        Err(_) => { println!("error"); }
    }
    workbook.close();
}

fn init_sheet(mut sheet: &mut Worksheet) {
    sheet.set_column(0, 100, 3.0, None);
}

fn create_format() -> Format {
    let mut format = Format::new();
    format.set_font_name("メイリオ");
    format.set_font_color(xlsxwriter::format::FormatColor::Black);
    return format; 
}

fn create_h1_format() -> Format {
    let mut format = create_format();
    format.set_font_size(20.0);
    return format;
}

fn create_h2_format() -> Format {
    let mut format  = create_format();
    format.set_font_size(16.0);
    return format;
}

fn create_h3_format() -> Format {
    let mut format  = create_format();
    format.set_font_size(13.0);
    return format;
}

fn create_bold_format() -> Format {
    let mut format  = create_format();
    format.set_bold();
    return format;
}

fn create_inline_format() -> Format {
    let mut format  = create_format();
    format.set_bg_color(xlsxwriter::format::FormatColor::Silver);
    return format; 
}

fn create_quotation_format() -> Format {
    let mut format  = create_format();
    format.set_bg_color(xlsxwriter::format::FormatColor::Silver);
    return format; 
}

fn create_code_format() -> Format {
    let mut format  = create_format();
    format.set_bg_color(xlsxwriter::format::FormatColor::Silver);
    return format;
}

fn insert_image(mut sheet: &mut Worksheet, image_path: &str, y: WorksheetRow) {
    let img_reader = image::io::Reader::open(image_path).unwrap();
    let img = img_reader.decode().unwrap();

    sheet.set_row(y, 230.0, None);
    let resized_img = img.resize_exact(600, 300, image::imageops::FilterType::Triangle);
    let mut bytes: Vec<u8> = Vec::new();
    resized_img.write_to(&mut Cursor::new(&mut bytes), image::ImageFormat::Png).unwrap();

    sheet.insert_image_buffer(y, 6, &bytes);
}

fn write_string(mut sheet: &mut Worksheet, text: &str, x: WorksheetCol, y: WorksheetRow, format: Option<Format>) {
    let newline_number = text.clone().split("\n").count();
    sheet.set_row(y, 28.0 + ((newline_number - 1) as f64) * 8.0, None);
    sheet.write_string(y, 6 + x, text, format.as_ref());
}