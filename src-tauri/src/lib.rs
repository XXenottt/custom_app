use calamine::{open_workbook_auto, Data, Reader};
use chrono::Local;
use pdf_extract::extract_text;
use regex::Regex;
use rust_xlsxwriter::{Color, Format, FormatBorder, Workbook};
use serde::Serialize;
use std::collections::HashMap;
use std::fs;
use std::path::{Path, PathBuf};

// ── Constants ────────────────────────────────────────────────────────────────

const SELLER: &str = "Huawei Device Co., Ltd.\nNo.2 of Xincheng Road,Songshan Lake Zone,Dongguan,Guangdong,P.R China";
const CONSIGNEE: &str = "HUAWEI DEVICE (HONG KONG) CO., LTD. C/O Huawei Technologies Netherlands B.V.\nLaan van Vredenoord 56, 2289DJ, Rijswijk, The Netherlands, VAT NL823106664B01/juxincheng/0031628291362/juxincheng@huawei.com";

// ── Data structures ──────────────────────────────────────────────────────────

#[derive(Debug, Clone)]
struct InvoiceLine {
    invoice_no: String,
    invoice_date: String,
    line_no: u32,
    part_no: String,
    description: String,
    country_of_origin: String,
    hs_code: String,
    qty: f64,
    uom: String,
    total_amount: f64,
    currency: String,
    trade_term: String,
    transport_mode: String,
    gw: f64,
    nw: f64,
    cases: u32,
}

#[derive(Debug, Default)]
struct PartWeight {
    total_gw: f64,
    total_nw: f64,
    total_qty: f64,
    total_cases: f64,
}

#[derive(Debug)]
struct MawbData {
    awb_no: String,
    total_gw: f64,
    total_cases: u32,
    total_freight: f64,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct ProcessResult {
    inv_files: usize,
    pl_files: usize,
    inv_rows: usize,
    pl_rows: usize,
    skipped_files: usize,
    output_path: String,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct GenerateResult {
    boe_files: Vec<String>,
    fiton_files: Vec<String>,
    inv_count: usize,
    line_count: usize,
    warnings: Vec<String>,
}

// ── Helpers ──────────────────────────────────────────────────────────────────

fn cell_str(range: &calamine::Range<Data>, row: u32, col: u32) -> String {
    match range.get_value((row, col)) {
        Some(Data::Empty) | None => String::new(),
        Some(v) => v.to_string().trim().to_string(),
    }
}

fn cell_f64(range: &calamine::Range<Data>, row: u32, col: u32) -> f64 {
    match range.get_value((row, col)) {
        Some(Data::Float(f)) => *f,
        Some(Data::Int(i)) => *i as f64,
        Some(Data::String(s)) => s.trim().parse().unwrap_or(0.0),
        _ => 0.0,
    }
}

fn is_numeric_cell(v: Option<&Data>) -> bool {
    match v {
        Some(Data::Float(_)) | Some(Data::Int(_)) => true,
        Some(Data::String(s)) => s.trim().parse::<f64>().is_ok(),
        _ => false,
    }
}

// ── Invoice parsing ──────────────────────────────────────────────────────────

fn parse_invoice(file: &Path) -> Result<Vec<InvoiceLine>, String> {
    let mut wb = open_workbook_auto(file).map_err(|e| e.to_string())?;
    let sname = wb.sheet_names().first().cloned().ok_or("no sheet")?;
    let range = wb.worksheet_range(&sname).map_err(|e| e.to_string())?;

    let invoice_no = cell_str(&range, 8, 8);
    let invoice_date = cell_str(&range, 9, 8);
    let trade_term = cell_str(&range, 11, 8);
    let currency_default = cell_str(&range, 14, 8);
    let transport = cell_str(&range, 15, 8);

    let mut lines = Vec::new();
    let mut row = 18u32; // row 19 (0-indexed)

    loop {
        let line_val = range.get_value((row, 0));
        if !is_numeric_cell(line_val) {
            break;
        }
        let line_no = cell_f64(&range, row, 0) as u32;
        let part_no = cell_str(&range, row, 1);
        let description = cell_str(&range, row, 2);
        let coo = cell_str(&range, row, 3);
        let hs_code = cell_str(&range, row, 4);
        let qty = cell_f64(&range, row, 5);
        let uom = cell_str(&range, row, 6);
        let total_amount = cell_f64(&range, row, 9);
        let line_currency = {
            let c = cell_str(&range, row, 10);
            if c.is_empty() { currency_default.clone() } else { c }
        };

        lines.push(InvoiceLine {
            invoice_no: invoice_no.clone(),
            invoice_date: invoice_date.clone(),
            line_no,
            part_no,
            description,
            country_of_origin: coo,
            hs_code,
            qty,
            uom,
            total_amount,
            currency: line_currency,
            trade_term: trade_term.clone(),
            transport_mode: transport.clone(),
            gw: 0.0,
            nw: 0.0,
            cases: 0,
        });
        row += 1;
    }

    Ok(lines)
}

// ── Packing list parsing ─────────────────────────────────────────────────────

fn parse_packing_lists(dir: &Path) -> HashMap<String, PartWeight> {
    let mut weights: HashMap<String, PartWeight> = HashMap::new();

    let entries = match fs::read_dir(dir) {
        Ok(e) => e,
        Err(_) => return weights,
    };

    for entry in entries.flatten() {
        let path = entry.path();
        let name = path.file_name().and_then(|n| n.to_str()).unwrap_or("");
        if !name.to_lowercase().contains("packing list") {
            continue;
        }
        let _ = parse_one_packing_list(&path, &mut weights);
    }

    weights
}

fn parse_one_packing_list(file: &Path, weights: &mut HashMap<String, PartWeight>) -> Result<(), String> {
    let mut wb = open_workbook_auto(file).map_err(|e| e.to_string())?;
    let sname = wb.sheet_names().first().cloned().ok_or("no sheet")?;
    let range = wb.worksheet_range(&sname).map_err(|e| e.to_string())?;

    // Data rows start at row 9 (0-indexed = 8), stop at "Total："
    let mut row = 8u32;
    loop {
        let a = cell_str(&range, row, 0);
        if a.starts_with("Total") || a.starts_with("Part:") {
            break;
        }
        let part_no = cell_str(&range, row, 1); // col B
        if part_no.is_empty() {
            row += 1;
            if row > 500 { break; }
            continue;
        }
        let qty_cartons = cell_f64(&range, row, 3);  // col D
        let qty_pcs = cell_f64(&range, row, 8);       // col I (total pcs in this row)
        let gw = cell_f64(&range, row, 10);            // col K
        let nw = cell_f64(&range, row, 11);            // col L

        let entry = weights.entry(part_no).or_default();
        entry.total_gw += gw;
        entry.total_nw += nw;
        entry.total_qty += qty_pcs;
        entry.total_cases += qty_cartons;

        row += 1;
    }

    Ok(())
}

// ── MAWB PDF parsing ─────────────────────────────────────────────────────────

fn parse_mawb(file: &Path) -> Result<MawbData, String> {
    let text = extract_text(file).map_err(|e| e.to_string())?;

    // AWB number: pattern like "784-83551123"
    let awb_re = Regex::new(r"\b(\d{3}-\d{8})\b").unwrap();
    let awb_no = awb_re.captures(&text)
        .and_then(|c| c.get(1))
        .map(|m| m.as_str().to_string())
        .unwrap_or_default();

    // Line like: "42 1751.0 K Q 1751.0 42.67 74715.17"
    // pieces  gw  unit class  chargeable  rate  total
    let freight_re = Regex::new(
        r"(\d+)\s+([\d.]+)\s+K\s+Q\s+[\d.]+\s+[\d.]+\s+([\d.]+)"
    ).unwrap();
    let (total_cases, total_gw, total_freight) = freight_re.captures(&text)
        .map(|c| (
            c[1].parse::<u32>().unwrap_or(0),
            c[2].parse::<f64>().unwrap_or(0.0),
            c[3].parse::<f64>().unwrap_or(0.0),
        ))
        .unwrap_or((0, 0.0, 0.0));

    Ok(MawbData { awb_no, total_gw, total_cases, total_freight })
}

// ── BOE generation ───────────────────────────────────────────────────────────

fn generate_boe(
    lines: &[InvoiceLine],
    mawb: &MawbData,
    output_path: &Path,
    currency: &str,
) -> Result<(), Box<dyn std::error::Error>> {
    let mut wb = Workbook::new();
    let ws = wb.add_worksheet().set_name("Sheet1")?;

    // Formats
    let hdr = Format::new()
        .set_bold()
        .set_background_color(Color::RGB(0xD9E1F2))
        .set_border(FormatBorder::Thin);
    let label = Format::new().set_bold();
    let wrap = Format::new().set_text_wrap();
    let num2 = Format::new().set_num_format("0.00");

    // ── Row 1: labels ──
    ws.write_with_format(0, 0, "Seller", &label)?;
    ws.write_with_format(0, 4, "Consignee", &label)?;
    ws.write_with_format(0, 9, "Trade Term", &label)?;
    ws.write_with_format(0, 10, "Total Gross Weight", &label)?;
    ws.write_with_format(0, 11, "Total Net Weight", &label)?;
    ws.write_with_format(0, 12, "Reference Freight (CNY)", &label)?;
    ws.write_with_format(0, 14, "Transport Mode", &label)?;
    ws.write_with_format(0, 15, "Departure Country/Region", &label)?;
    ws.write_with_format(0, 16, "Destination Country/Region", &label)?;

    // ── Row 2: values ──
    let total_gw: f64 = lines.iter().map(|l| l.gw).sum();
    let total_nw: f64 = lines.iter().map(|l| l.nw).sum();
    let freight = if mawb.total_gw > 0.0 {
        (mawb.total_freight * total_gw / mawb.total_gw * 100.0).round() / 100.0
    } else {
        0.0
    };
    let trade_term = lines.first().map(|l| l.trade_term.as_str()).unwrap_or("");
    let transport = lines.first().map(|l| l.transport_mode.as_str()).unwrap_or("BY AIR");

    ws.write_with_format(1, 0, SELLER, &wrap)?;
    ws.write_with_format(1, 4, CONSIGNEE, &wrap)?;
    ws.write_with_format(1, 9, trade_term, &wrap)?;
    ws.write_with_format(1, 10, total_gw, &num2)?;
    ws.write_with_format(1, 11, total_nw, &num2)?;
    ws.write_with_format(1, 12, freight, &num2)?;
    ws.write_with_format(1, 14, transport, &Default::default())?;
    ws.write_with_format(1, 15, "CN", &Default::default())?;
    ws.write_with_format(1, 16, "NL", &Default::default())?;

    // ── Row 4: column headers ──
    let col_headers = [
        "Invoice No.", "Invoice Date", "Line No.", "HS Code", "C.O.O.",
        "Item Code", "Gross Weight", "Net Weight", "Item Desc", "Item Quantity",
        "Unit", "Total Hs Code Gross weight", "Total Hs Code Net weight",
        "HS Code", "Amount", "Total Hscode Amount", "Total Amount", "BOE No.",
    ];
    for (c, h) in col_headers.iter().enumerate() {
        ws.write_with_format(3, c as u16, *h, &hdr)?;
    }

    // ── Sort lines by HS code, then invoice, then line_no ──
    let mut sorted = lines.to_vec();
    sorted.sort_by(|a, b| {
        a.hs_code.cmp(&b.hs_code)
            .then(a.invoice_no.cmp(&b.invoice_no))
            .then(a.line_no.cmp(&b.line_no))
    });

    // Pre-compute per-HS-code totals
    let mut hs_gw: HashMap<String, f64> = HashMap::new();
    let mut hs_nw: HashMap<String, f64> = HashMap::new();
    let mut hs_amount: HashMap<String, f64> = HashMap::new();
    for l in &sorted {
        *hs_gw.entry(l.hs_code.clone()).or_default() += l.gw;
        *hs_nw.entry(l.hs_code.clone()).or_default() += l.nw;
        *hs_amount.entry(l.hs_code.clone()).or_default() += l.total_amount;
    }

    let total_amount: f64 = sorted.iter().map(|l| l.total_amount).sum();

    // ── Write line items ──
    let mut current_hs = String::new();
    let mut first_row = true;
    for (i, line) in sorted.iter().enumerate() {
        let r = (4 + i) as u32;
        let is_new_hs = line.hs_code != current_hs;
        if is_new_hs {
            current_hs = line.hs_code.clone();
        }

        ws.write(r, 0, line.invoice_no.as_str())?;
        ws.write(r, 1, line.invoice_date.as_str())?;
        ws.write(r, 2, line.line_no)?;
        ws.write(r, 3, line.hs_code.as_str())?;
        ws.write(r, 4, line.country_of_origin.as_str())?;
        ws.write(r, 5, line.part_no.as_str())?;
        ws.write_with_format(r, 6, line.gw, &num2)?;
        ws.write_with_format(r, 7, line.nw, &num2)?;
        ws.write(r, 8, line.description.as_str())?;
        ws.write(r, 9, line.qty)?;
        ws.write(r, 10, line.uom.as_str())?;

        if is_new_hs {
            ws.write_with_format(r, 11, *hs_gw.get(&line.hs_code).unwrap_or(&0.0), &num2)?;
            ws.write_with_format(r, 12, *hs_nw.get(&line.hs_code).unwrap_or(&0.0), &num2)?;
            ws.write(r, 13, line.hs_code.as_str())?;
            let amt = hs_amount.get(&line.hs_code).unwrap_or(&0.0);
            ws.write(r, 14, format!("{:.2} {}", line.total_amount, currency).as_str())?;
            ws.write_with_format(r, 15, *amt, &num2)?;
        } else {
            ws.write(r, 14, format!("{:.2} {}", line.total_amount, currency).as_str())?;
        }

        if first_row {
            ws.write_with_format(r, 16, total_amount, &num2)?;
            ws.write(r, 17, "")?; // BOE No. - left blank
            first_row = false;
        }
    }

    // Column widths
    ws.set_column_width(0, 22)?;
    ws.set_column_width(4, 18)?;
    ws.set_column_width(8, 30)?;
    ws.set_column_width(14, 20)?;

    wb.save(output_path)?;
    Ok(())
}

// ── Fiton generation ─────────────────────────────────────────────────────────

fn generate_fiton(
    lines: &[InvoiceLine],
    output_path: &Path,
    currency: &str,
) -> Result<(), Box<dyn std::error::Error>> {
    let mut wb = Workbook::new();
    let ws = wb.add_worksheet().set_name("Sheet1")?;

    let hdr = Format::new()
        .set_bold()
        .set_background_color(Color::RGB(0xFFFF00))
        .set_border(FormatBorder::Thin);
    let red_hdr = Format::new()
        .set_bold()
        .set_background_color(Color::RGB(0xFF0000))
        .set_font_color(Color::White)
        .set_border(FormatBorder::Thin);
    let num2 = Format::new().set_num_format("0.00");

    // Column headers (field codes from Fiton requirement)
    let headers: &[(&str, bool)] = &[
        ("A",  true),  // Artikelnummer
        ("C",  true),  // Goederencode (HS code)
        ("F",  true),  // Goederenomschrijving
        ("G",  true),  // Verpakking
        ("H",  true),  // Aantal colli
        ("I",  true),  // Merken en nummers
        ("J",  true),  // Bruto massa
        ("K",  true),  // Netto massa
        ("L",  true),  // Aanvullende eenheden
        ("M",  false), // Gevraagde regeling
        ("N",  false), // Voorafgaande regeling
        ("P",  false), // Communautaire preferentie
        ("Q",  true),  // Land van oorsprong
        ("T",  false), // Waarderingsmethode
        ("U",  false), // Waarderingsindicator
        ("AD", true),  // Prijs van de goederen
        ("AE", true),  // Valuta prijs
        ("AH", false), // Valuta kosten
        ("BO", false), // Aanvullende informatie
        ("BP", false), // Beschrijving aanvullende informatie
        ("BQ", false), // Aanvullende referentie
        ("CA", true),  // Bewijsstuk
        ("CB", true),  // Nummer bewijsstuk
    ];

    for (c, (name, is_auto)) in headers.iter().enumerate() {
        let fmt = if *is_auto { &red_hdr } else { &hdr };
        ws.write_with_format(0, c as u16, *name, fmt)?;
    }

    // Group lines by HS code
    let mut hs_order: Vec<String> = Vec::new();
    let mut hs_groups: HashMap<String, Vec<&InvoiceLine>> = HashMap::new();
    for line in lines {
        if !hs_groups.contains_key(&line.hs_code) {
            hs_order.push(line.hs_code.clone());
        }
        hs_groups.entry(line.hs_code.clone()).or_default().push(line);
    }
    hs_order.sort();

    for (seq, hs) in hs_order.iter().enumerate() {
        let group = &hs_groups[hs];
        let r = (1 + seq) as u32;

        let total_gw: f64 = group.iter().map(|l| l.gw).sum();
        let total_nw: f64 = group.iter().map(|l| l.nw).sum();
        let total_qty: f64 = group.iter().map(|l| l.qty).sum();
        let total_amount: f64 = group.iter().map(|l| l.total_amount).sum();
        let total_cases: u32 = group.iter().map(|l| l.cases).sum();
        let description = group.first().map(|l| l.description.as_str()).unwrap_or("");
        let coo = group.first().map(|l| l.country_of_origin.as_str()).unwrap_or("CN");
        // Collect unique invoice numbers
        let mut inv_nos: Vec<String> = group.iter().map(|l| l.invoice_no.clone()).collect();
        inv_nos.dedup();
        let inv_str = inv_nos.join("; ");

        // A: seq number
        ws.write(r, 0, (seq + 1) as u32)?;
        // C: HS code
        ws.write(r, 1, hs.as_str())?;
        // F: description
        ws.write(r, 2, description)?;
        // G: packaging type
        ws.write(r, 3, "CT")?;
        // H: number of cartons
        ws.write(r, 4, total_cases)?;
        // I: marks and numbers (invoice nos)
        ws.write(r, 5, inv_str.as_str())?;
        // J: gross weight
        ws.write_with_format(r, 6, total_gw, &num2)?;
        // K: net weight
        ws.write_with_format(r, 7, total_nw, &num2)?;
        // L: supplementary units (qty)
        ws.write(r, 8, total_qty)?;
        // M: Gevraagde regeling
        ws.write(r, 9, "40")?;
        // N: Voorafgaande regeling
        ws.write(r, 10, "71")?;
        // P: Communautaire preferentie
        ws.write(r, 11, "100")?;
        // Q: country of origin
        ws.write(r, 12, if coo.to_lowercase().contains("china") { "CN" } else { coo })?;
        // T: Waarderingsmethode
        ws.write(r, 13, "1")?;
        // U: Waarderingsindicator
        ws.write(r, 14, "0000")?;
        // AD: price
        ws.write_with_format(r, 15, total_amount, &num2)?;
        // AE: currency
        ws.write(r, 16, currency)?;
        // AH: blank (no cost currency specified)
        ws.write(r, 17, "")?;
        // BO, BP: blank (Extron-specific, not applicable here)
        ws.write(r, 18, "")?;
        ws.write(r, 19, "")?;
        // BQ: blank (manual input, customs-specific)
        ws.write(r, 20, "")?;
        // CA: document type
        ws.write(r, 21, "N935")?;
        // CB: invoice number
        ws.write(r, 22, inv_str.as_str())?;
    }

    // Column widths
    ws.set_column_width(1, 14)?;
    ws.set_column_width(2, 35)?;
    ws.set_column_width(5, 25)?;
    ws.set_column_width(22, 25)?;

    wb.save(output_path)?;
    Ok(())
}

// ── Tauri commands ───────────────────────────────────────────────────────────

#[tauri::command]
fn generate_customs_docs(
    input_dir: String,
    mawb_path: String,
    output_dir: String,
) -> Result<GenerateResult, String> {
    let input = PathBuf::from(input_dir.trim());
    if !input.is_dir() {
        return Err("输入目录不存在".to_string());
    }
    let out_dir = PathBuf::from(output_dir.trim());
    if !out_dir.is_dir() {
        return Err("输出目录不存在".to_string());
    }
    let mawb_file = PathBuf::from(mawb_path.trim());

    let mut warnings: Vec<String> = Vec::new();

    // Parse MAWB
    let mawb = if mawb_file.exists() {
        parse_mawb(&mawb_file).unwrap_or_else(|e| {
            warnings.push(format!("MAWB 解析失败: {e}"));
            MawbData { awb_no: String::new(), total_gw: 0.0, total_cases: 0, total_freight: 0.0 }
        })
    } else {
        warnings.push("未找到 MAWB PDF，重量将按比例估算".to_string());
        MawbData { awb_no: String::new(), total_gw: 0.0, total_cases: 0, total_freight: 0.0 }
    };

    // Parse packing lists
    let weights = parse_packing_lists(&input);

    // Parse all invoices
    let entries = fs::read_dir(&input).map_err(|e| format!("读取目录失败: {e}"))?;
    let mut all_lines: Vec<InvoiceLine> = Vec::new();
    let mut inv_count = 0usize;

    for entry in entries.flatten() {
        let path = entry.path();
        let name = path.file_name().and_then(|n| n.to_str()).unwrap_or("");
        if name.contains("INV") && name.ends_with(".xlsx") {
            match parse_invoice(&path) {
                Ok(lines) => {
                    inv_count += 1;
                    all_lines.extend(lines);
                }
                Err(e) => warnings.push(format!("跳过 {name}: {e}")),
            }
        }
    }

    let line_count = all_lines.len();

    // Allocate weights
    allocate_weights(&mut all_lines, &weights, &mawb, &mut warnings);

    // Split by currency
    let hkd_lines: Vec<InvoiceLine> = all_lines.iter().filter(|l| l.currency == "HKD").cloned().collect();
    let eur_lines: Vec<InvoiceLine> = all_lines.iter().filter(|l| l.currency == "EUR").cloned().collect();

    let date_str = Local::now().format("%Y%m%d").to_string();
    let mut boe_files = Vec::new();
    let mut fiton_files = Vec::new();

    if !hkd_lines.is_empty() {
        let boe_path = out_dir.join(format!("BOE_HKD_{date_str}.xlsx"));
        let fiton_path = out_dir.join(format!("Fiton_HKD_{date_str}.xlsx"));
        generate_boe(&hkd_lines, &mawb, &boe_path, "HKD")
            .map_err(|e| format!("BOE HKD 生成失败: {e}"))?;
        generate_fiton(&hkd_lines, &fiton_path, "HKD")
            .map_err(|e| format!("Fiton HKD 生成失败: {e}"))?;
        boe_files.push(boe_path.to_string_lossy().to_string());
        fiton_files.push(fiton_path.to_string_lossy().to_string());
    }

    if !eur_lines.is_empty() {
        let boe_path = out_dir.join(format!("BOE_EUR_{date_str}.xlsx"));
        let fiton_path = out_dir.join(format!("Fiton_EUR_{date_str}.xlsx"));
        generate_boe(&eur_lines, &mawb, &boe_path, "EUR")
            .map_err(|e| format!("BOE EUR 生成失败: {e}"))?;
        generate_fiton(&eur_lines, &fiton_path, "EUR")
            .map_err(|e| format!("Fiton EUR 生成失败: {e}"))?;
        boe_files.push(boe_path.to_string_lossy().to_string());
        fiton_files.push(fiton_path.to_string_lossy().to_string());
    }

    Ok(GenerateResult { boe_files, fiton_files, inv_count, line_count, warnings })
}

#[tauri::command]
fn process_excel_files(input_dir: String, output_path: String) -> Result<ProcessResult, String> {
    let input = PathBuf::from(input_dir.trim());
    if !input.exists() || !input.is_dir() {
        return Err("输入目录不存在或不是目录".to_string());
    }
    let output = PathBuf::from(output_path.trim());
    if output.as_os_str().is_empty() {
        return Err("输出文件路径不能为空".to_string());
    }
    if let Some(parent) = output.parent() {
        if !parent.as_os_str().is_empty() && !parent.exists() {
            return Err("输出目录不存在".to_string());
        }
    }

    let mut files = Vec::new();
    for entry in fs::read_dir(&input).map_err(|e| format!("读取目录失败: {e}"))?.flatten() {
        let path = entry.path();
        if path.is_file() { files.push(path); }
    }
    files.sort();

    let inv_files: Vec<PathBuf> = files.iter()
        .filter(|p| p.file_name().and_then(|v| v.to_str()).map(|s| s.contains("INV")).unwrap_or(false))
        .cloned().collect();
    let pl_files: Vec<PathBuf> = files.iter()
        .filter(|p| p.file_name().and_then(|v| v.to_str()).map(|s| s.to_lowercase().contains("packing list")).unwrap_or(false))
        .cloned().collect();

    let mut inv_data: Vec<Vec<String>> = Vec::new();
    let mut pl_data: Vec<Vec<String>> = Vec::new();
    let mut skipped_files = 0usize;

    for file in &inv_files {
        match process_inv_file(file) {
            Ok(rows) => inv_data.extend(rows),
            Err(_) => skipped_files += 1,
        }
    }

    let cases_re = Regex::new(r"Total：(\d+)\s+CASES").map_err(|e| format!("正则初始化失败: {e}"))?;
    let net_re = Regex::new(r"net weight\(KG\)：([\d.]+)").map_err(|e| format!("正则初始化失败: {e}"))?;

    for file in &pl_files {
        match process_packing_list_file(file, &cases_re, &net_re) {
            Ok(Some(row)) => pl_data.push(row),
            Ok(None) => {}
            Err(_) => skipped_files += 1,
        }
    }

    write_output_file(&output, &inv_data, &pl_data).map_err(|e| format!("写入输出文件失败: {e}"))?;

    Ok(ProcessResult {
        inv_files: inv_files.len(),
        pl_files: pl_files.len(),
        inv_rows: inv_data.len(),
        pl_rows: pl_data.len(),
        skipped_files,
        output_path: output.to_string_lossy().to_string(),
    })
}

fn allocate_weights(
    lines: &mut Vec<InvoiceLine>,
    weights: &HashMap<String, PartWeight>,
    mawb: &MawbData,
    warnings: &mut Vec<String>,
) {
    // Total qty across all lines (for fallback proportional allocation)
    let total_qty: f64 = lines.iter().map(|l| l.qty).sum();

    for line in lines.iter_mut() {
        if let Some(pw) = weights.get(&line.part_no) {
            if pw.total_qty > 0.0 {
                let ratio = line.qty / pw.total_qty;
                line.gw = (pw.total_gw * ratio * 100.0).round() / 100.0;
                line.nw = (pw.total_nw * ratio * 100.0).round() / 100.0;
                line.cases = (pw.total_cases * ratio).round() as u32;
            }
        } else {
            // Fallback: allocate proportionally from MAWB total
            if total_qty > 0.0 {
                let ratio = line.qty / total_qty;
                line.gw = (mawb.total_gw * ratio * 100.0).round() / 100.0;
                line.nw = (mawb.total_gw * ratio * 0.85 * 100.0).round() / 100.0;
                line.cases = (mawb.total_cases as f64 * ratio).round() as u32;
                warnings.push(format!(
                    "Part {} not found in packing lists, weight estimated proportionally",
                    line.part_no
                ));
            }
        }
    }
}

// ── Legacy helpers (used by process_excel_files) ─────────────────────────────

fn process_inv_file(file: &Path) -> Result<Vec<Vec<String>>, String> {
    let mut workbook = open_workbook_auto(file).map_err(|e| e.to_string())?;
    let sheet_name = workbook.sheet_names().first().cloned().ok_or_else(|| "找不到工作表".to_string())?;
    let range = workbook.worksheet_range(&sheet_name).map_err(|e| e.to_string())?;
    let ijk9_value = cell_str(&range, 8, 8);
    let mut rows = Vec::new();
    let mut row_index = 18usize;
    let mut first_row = true;
    while is_numeric_cell(range.get_value((row_index as u32, 0))) {
        let mut row_values = Vec::new();
        for col in 1..=6u32 {
            row_values.push(cell_str(&range, row_index as u32, col));
        }
        row_values.push(cell_str(&range, row_index as u32, 10));
        for col in 8..=9u32 {
            row_values.push(cell_str(&range, row_index as u32, col));
        }
        let mut final_row = Vec::new();
        final_row.push(if first_row { ijk9_value.clone() } else { String::new() });
        final_row.push(String::new());
        final_row.extend(row_values);
        rows.push(final_row);
        first_row = false;
        row_index += 1;
    }
    Ok(rows)
}

fn process_packing_list_file(
    file: &Path,
    cases_re: &Regex,
    net_re: &Regex,
) -> Result<Option<Vec<String>>, String> {
    let mut workbook = open_workbook_auto(file).map_err(|e| e.to_string())?;
    let sheet_name = workbook.sheet_names().first().cloned().ok_or_else(|| "找不到工作表".to_string())?;
    let range = workbook.worksheet_range(&sheet_name).map_err(|e| e.to_string())?;
    for row in range.rows() {
        for cell in row {
            let text = match cell {
                Data::String(s) => s.clone(),
                _ => continue,
            };
            if text.contains("Total：") && text.contains("net weight(KG)：") {
                if let (Some(cases), Some(net_weight)) = (cases_re.captures(&text), net_re.captures(&text)) {
                    return Ok(Some(vec![
                        net_weight.get(1).map(|v| v.as_str().to_string()).unwrap_or_default(),
                        cases.get(1).map(|v| v.as_str().to_string()).unwrap_or_default(),
                    ]));
                }
            }
        }
    }
    Ok(None)
}

fn write_output_file(
    output: &Path,
    inv_data: &[Vec<String>],
    pl_data: &[Vec<String>],
) -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook = Workbook::new();
    if inv_data.is_empty() && pl_data.is_empty() {
        let sheet = workbook.add_worksheet().set_name("Sheet1")?;
        sheet.write_string(0, 0, "No matching data found")?;
    } else {
        if !inv_data.is_empty() {
            let sheet1 = workbook.add_worksheet().set_name("Sheet1")?;
            for (r, row) in inv_data.iter().enumerate() {
                for (c, value) in row.iter().enumerate() {
                    sheet1.write_string(r as u32, c as u16, value)?;
                }
            }
        }
        if !pl_data.is_empty() {
            let sheet2 = workbook.add_worksheet().set_name("Sheet2")?;
            for (r, row) in pl_data.iter().enumerate() {
                for (c, value) in row.iter().enumerate() {
                    sheet2.write_string(r as u32, c as u16, value)?;
                }
            }
        }
    }
    workbook.save(output)?;
    Ok(())
}

// ── App entry point ───────────────────────────────────────────────────────────

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_opener::init())
        .invoke_handler(tauri::generate_handler![
            process_excel_files,
            generate_customs_docs,
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}