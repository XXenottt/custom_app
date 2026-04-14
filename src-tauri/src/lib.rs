use calamine::{open_workbook_auto, Data, Reader};
use regex::Regex;
use rust_xlsxwriter::Workbook;
use serde::Serialize;
use std::fs;
use std::path::{Path, PathBuf};

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
    let entries = fs::read_dir(&input).map_err(|e| format!("读取目录失败: {e}"))?;
    for entry in entries {
        let entry = match entry {
            Ok(v) => v,
            Err(_) => continue,
        };
        let path = entry.path();
        if path.is_file() {
            files.push(path);
        }
    }
    files.sort();

    let inv_files: Vec<PathBuf> = files
        .iter()
        .filter(|p| {
            p.file_name()
                .and_then(|v| v.to_str())
                .map(|s| s.contains("INV"))
                .unwrap_or(false)
        })
        .cloned()
        .collect();

    let pl_files: Vec<PathBuf> = files
        .iter()
        .filter(|p| {
            p.file_name()
                .and_then(|v| v.to_str())
                .map(|s| s.to_lowercase().contains("packing list"))
                .unwrap_or(false)
        })
        .cloned()
        .collect();

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

fn process_inv_file(file: &Path) -> Result<Vec<Vec<String>>, String> {
    let mut workbook = open_workbook_auto(file).map_err(|e| e.to_string())?;
    let sheet_name = workbook
        .sheet_names()
        .first()
        .cloned()
        .ok_or_else(|| "找不到工作表".to_string())?;
    let range = workbook
        .worksheet_range(&sheet_name)
        .map_err(|e| e.to_string())?;

    let ijk9_value = cell_to_string(range.get_value((8, 8)));

    let mut rows = Vec::new();
    let mut row_index = 18usize;
    let mut first_row = true;

    while has_value(range.get_value((row_index as u32, 0))) {
        let mut row_values = Vec::new();

        for col in 1..=6 {
            row_values.push(cell_to_string(range.get_value((row_index as u32, col))));
        }

        row_values.push(cell_to_string(range.get_value((row_index as u32, 10))));

        for col in 8..=9 {
            row_values.push(cell_to_string(range.get_value((row_index as u32, col))));
        }

        let mut final_row = Vec::new();
        final_row.push(if first_row {
            ijk9_value.clone()
        } else {
            String::new()
        });
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
    let sheet_name = workbook
        .sheet_names()
        .first()
        .cloned()
        .ok_or_else(|| "找不到工作表".to_string())?;
    let range = workbook
        .worksheet_range(&sheet_name)
        .map_err(|e| e.to_string())?;

    for row in range.rows() {
        for cell in row {
            let text = cell_to_string(Some(cell));
            if text.contains("Total：") && text.contains("net weight(KG)：") {
                if let (Some(cases), Some(net_weight)) = (cases_re.captures(&text), net_re.captures(&text)) {
                    return Ok(Some(vec![
                        net_weight
                            .get(1)
                            .map(|v| v.as_str().to_string())
                            .unwrap_or_default(),
                        cases
                            .get(1)
                            .map(|v| v.as_str().to_string())
                            .unwrap_or_default(),
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

fn has_value(cell: Option<&Data>) -> bool {
    cell.map(|v| !matches!(v, Data::Empty)).unwrap_or(false)
}

fn cell_to_string(cell: Option<&Data>) -> String {
    match cell {
        Some(Data::Empty) | None => String::new(),
        Some(value) => value.to_string(),
    }
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_opener::init())
        .invoke_handler(tauri::generate_handler![process_excel_files])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
