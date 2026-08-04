import os
import pandas as pd
import datetime
import unicodedata
import sys
import json
import shutil
import glob

# Ensure stdout uses UTF-8 to prevent console encoding issues on Windows
try:
    sys.stdout.reconfigure(encoding='utf-8')
except AttributeError:
    pass

raw_data_dir = r"d:\CJ\11. Adhoc\T8\WW\raw_data"
js_output = r"d:\CJ\11. Adhoc\T8\WW\coaching_data.js"
xlsx_output = r"d:\CJ\11. Adhoc\T8\WW\Bao_Cao_Coaching_GSBH.xlsx"

raw_data_abs = os.path.abspath(raw_data_dir)
js_abs = os.path.abspath(js_output)
xlsx_abs = os.path.abspath(xlsx_output)

def clean_val(val):
    if val is None:
        return None
    if hasattr(val, 'year') and hasattr(val, 'month') and hasattr(val, 'day'):
        try:
            return datetime.date(val.year, val.month, val.day)
        except Exception:
            return str(val)
    return val

def normalize_text(text):
    if not isinstance(text, str):
        return ""
    return unicodedata.normalize('NFC', text.strip())

def get_week_of_month(date_val):
    if not isinstance(date_val, (datetime.date, datetime.datetime)):
        return 'Khác'
    day = date_val.day
    if day <= 11:
        return 'Tuần 1'
    elif day <= 18:
        return 'Tuần 2'
    elif day <= 25:
        return 'Tuần 3'
    else:
        return 'Tuần 4'

def get_target_day_count_by_date(date_val):
    if not isinstance(date_val, (datetime.date, datetime.datetime)):
        return 16.0
    day = date_val.day
    if day <= 11:
        return 4.0
    elif day <= 18:
        return 8.0
    elif day <= 25:
        return 12.0
    else:
        return 16.0

def standardize_code(code_val):
    code_str = str(code_val).strip().upper()
    if code_str.startswith('0') and code_str[1:].isdigit():
        return code_str.lstrip('0')
    if code_str.isdigit():
        return code_str.lstrip('0')
    return code_str

def assign_region(vung):
    vung_clean = normalize_text(vung).lower()
    if 'hà nội' in vung_clean or 'hải phòng' in vung_clean or 'đà nẵng' in vung_clean:
        return 'Miền Bắc'
    elif 'nha trang' in vung_clean or 'hcm' in vung_clean or 'cần thơ' in vung_clean:
        return 'Miền Nam'
    return 'Khác'

try:
    # Scan raw_data folder for files
    file_pattern_xlsb = os.path.join(raw_data_abs, "*.xlsb")
    file_pattern_xlsx = os.path.join(raw_data_abs, "*.xlsx")
    raw_files = glob.glob(file_pattern_xlsb) + glob.glob(file_pattern_xlsx)
    # Remove files that are open/locked (temporary start with ~$ )
    raw_files = [f for f in raw_files if not os.path.basename(f).startswith("~$")]
    
    if len(raw_files) == 0:
        print(f"Error: Không tìm thấy bất kỳ file .xlsb hay .xlsx nào trong thư mục {raw_data_abs}!")
        exit(1)
        
    all_dfs = []
    
    # Check if win32com is available (Windows local environment with Excel installed)
    use_excel_com = False
    if os.name == 'nt':
        try:
            import win32com.client as win32
            use_excel_com = True
        except Exception:
            pass
            
    if use_excel_com:
        print("Đang khởi tạo Excel COM để đọc dữ liệu thô (Windows local)...")
        try:
            os.system("taskkill /f /im excel.exe >nul 2>&1")
        except:
            pass
        
        excel = win32.Dispatch('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        
        for f_path in raw_files:
            print(f"Đang đọc file dữ liệu thô: {os.path.basename(f_path)}...")
            try:
                wb = excel.Workbooks.Open(os.path.abspath(f_path))
                sheet = wb.Sheets(1)
                raw_data = sheet.UsedRange.Value
                wb.Close(SaveChanges=False)
                
                if not raw_data or len(raw_data) < 5:
                    print(f"Bỏ qua file {os.path.basename(f_path)} vì không đủ dòng.")
                    continue
                    
                headers = [normalize_text(h) if h is not None else f"Col_{i}" for i, h in enumerate(raw_data[4])]
                
                data_rows = []
                for row in raw_data[5:]:
                    cleaned_row = [clean_val(cell) for cell in row]
                    data_rows.append(cleaned_row)
                    
                df_single = pd.DataFrame(data_rows, columns=headers)
                all_dfs.append(df_single)
                print(f"Đã đọc xong {len(df_single)} hàng dữ liệu từ {os.path.basename(f_path)}.")
            except Exception as file_err:
                print(f"Lỗi khi đọc file {os.path.basename(f_path)}: {file_err}")
                
        excel.Quit()
        excel = None
    else:
        print("Không tìm thấy Excel COM hoặc chạy trên Linux. Chuyển sang đọc bằng pandas cross-platform...")
        for f_path in raw_files:
            print(f"Đang đọc file dữ liệu thô bằng pandas: {os.path.basename(f_path)}...")
            try:
                ext = os.path.splitext(f_path)[1].lower()
                if ext == '.xlsb':
                    df_single = pd.read_excel(f_path, sheet_name=0, header=4, engine='pyxlsb')
                else:
                    df_single = pd.read_excel(f_path, sheet_name=0, header=4, engine='openpyxl')
                
                df_single.columns = [normalize_text(c) if c is not None else f"Col_{i}" for i, c in enumerate(df_single.columns)]
                
                # Convert Excel date numbers (float) to datetime
                if 'Từ Ngày' in df_single.columns:
                    def convert_excel_date(val):
                        if isinstance(val, (int, float)) and not pd.isna(val):
                            return (pd.to_datetime('1899-12-30') + pd.to_timedelta(val, unit='D')).date()
                        if isinstance(val, str) and val.strip():
                            try:
                                return pd.to_datetime(val.strip()).date()
                            except:
                                return val
                        if hasattr(val, 'date'):
                            return val.date()
                        return val
                    df_single['Từ Ngày'] = df_single['Từ Ngày'].apply(convert_excel_date)
                
                for col in df_single.columns:
                    df_single[col] = df_single[col].apply(lambda x: None if pd.isna(x) else x)
                    
                all_dfs.append(df_single)
                print(f"Đã đọc xong {len(df_single)} hàng dữ liệu từ {os.path.basename(f_path)}.")
            except Exception as file_err:
                print(f"Lỗi khi đọc file {os.path.basename(f_path)}: {file_err}")
                
    if len(all_dfs) == 0:
        print("Không có dữ liệu thô nào được load!")
        exit(1)
        
    df = pd.concat(all_dfs, ignore_index=True)
    print(f"Tổng số dữ liệu thô load được từ tất cả các file: {len(df)} dòng.")
    
    # 3. Standardize column contents
    text_cols = ['Tên Vùng', 'Tên NV Đăng Ký', 'Loại Công Việc', 'Trạng thái công việc', 'Trạng Thái Công Việc', 'Mã NV Đăng Ký']
    # Merge 'Trạng thái công việc' if present
    if 'Trạng thái công việc' in df.columns and 'Trạng Thái Công Việc' not in df.columns:
        df['Trạng Thái Công Việc'] = df['Trạng thái công việc']
        
    for col in text_cols:
        if col in df.columns:
            if col == 'Mã NV Đăng Ký':
                df[col] = df[col].apply(standardize_code)
            else:
                df[col] = df[col].astype(str).apply(normalize_text)
            
    # 4. Read master roster from DS SGBH.xlsx
    master_path = r"d:\CJ\11. Adhoc\T8\WW\DS SGBH.xlsx"
    master_abs = os.path.abspath(master_path)
    
    print("Đang đọc danh sách GSBH master...")
    df_master = pd.read_excel(master_abs)
    for col in df_master.columns:
        df_master[col] = df_master[col].astype(str).apply(normalize_text)
        
    df_master['Mã GSBH'] = df_master['Mã GSBH'].apply(standardize_code)
    df_master['Tên GSBH'] = df_master['Tên GSBH'].str.strip()
    
    def clean_mien(m):
        m_clean = m.upper()
        if 'BẮC' in m_clean:
            return 'Miền Bắc'
        elif 'NAM' in m_clean:
            return 'Miền Nam'
        return 'Khác'
    df_master['Miền'] = df_master['Tên miền'].apply(clean_mien)
    
    def clean_vung(v):
        v_clean = v.lower()
        if 'hà nội' in v_clean or 'hà nội' in v_clean:
            return 'Hà Nội'
        elif 'hải phòng' in v_clean or 'hải phòng' in v_clean:
            return 'Hải Phòng'
        elif 'đà nẵng' in v_clean:
            return 'Đà Nẵng'
        elif 'cần thơ' in v_clean:
            return 'Cần Thơ'
        elif 'hcm' in v_clean or 'hồ chí minh' in v_clean:
            return 'HCM'
        elif 'nha trang' in v_clean:
            return 'Nha Trang'
        return v
    df_master['Tên Vùng Sạch'] = df_master['Tên Vùng'].apply(clean_vung)
    
    gsbh_list = df_master.groupby('Mã GSBH').agg({
        'Miền': 'first',
        'Tên Vùng Sạch': 'first',
        'Tên GSBH': 'first'
    }).reset_index()
    
    gsbh_list = gsbh_list.rename(columns={
        'Mã GSBH': 'Mã NV Đăng Ký',
        'Tên GSBH': 'Tên NV Đăng Ký',
        'Tên Vùng Sạch': 'Tên Vùng'
    })
    gsbh_list = gsbh_list.sort_values(by=['Miền', 'Tên Vùng', 'Tên NV Đăng Ký']).reset_index(drop=True)
    print(f"Đã tải danh sách gồm {len(gsbh_list)} GSBH master từ file DS SGBH.xlsx.")
    
    # 5. Filter for coaching activities only ("Huấn Luyện")
    df_hl = df[df['Loại Công Việc'].str.lower() == 'huấn luyện'].copy()
    
    # Map Week for Coaching records (4 weeks only)
    df_hl['Tuần'] = df_hl['Từ Ngày'].apply(get_week_of_month)
    df_hl['Từ Ngày'] = pd.to_datetime(df_hl['Từ Ngày']).dt.date
    
    # Filter for valid statuses: Đăng ký (Đã Duyệt + Hoàn Thành), Hoàn thành (Hoàn Thành)
    valid_statuses = ['đã duyệt', 'hoàn thành']
    df_hl_valid = df_hl[df_hl['Trạng Thái Công Việc'].str.lower().isin(valid_statuses)].copy()
    
    # Create indicators
    df_hl_valid['Is_Registered'] = 1
    df_hl_valid['Is_Completed'] = df_hl_valid['Trạng Thái Công Việc'].str.lower().apply(lambda x: 1 if x == 'hoàn thành' else 0)
    
    # Merge master region/zone details back to df_hl_valid
    for c in ['Miền', 'Tên Vùng']:
        if c in df_hl_valid.columns:
            df_hl_valid = df_hl_valid.drop(columns=[c])
    df_hl_valid = df_hl_valid.merge(
        gsbh_list[['Mã NV Đăng Ký', 'Miền', 'Tên Vùng']], 
        on='Mã NV Đăng Ký', 
        how='left'
    )
    df_hl_valid['Miền'] = df_hl_valid['Miền'].fillna('Khác')
    df_hl_valid['Tên Vùng'] = df_hl_valid['Tên Vùng'].fillna('Khác')
     # 5.1 Extract Months list
    df_hl_valid['Tháng'] = pd.to_datetime(df_hl_valid['Từ Ngày']).apply(lambda x: x.strftime('%Y-%m') if hasattr(x, 'strftime') else str(x)[:7])
    available_months = sorted(df_hl_valid['Tháng'].unique())
    print(f"Các tháng tìm thấy trong dữ liệu: {available_months}")
    
    months_data = {}
    weeks = ['Tuần 1', 'Tuần 2', 'Tuần 3', 'Tuần 4']
    
    for cur_month in available_months:
        print(f"[*] Đang tính toán dữ liệu cho Tháng {cur_month}...")
        df_m = df_hl_valid[df_hl_valid['Tháng'] == cur_month].copy()
        
        # 6. Calculate weekly metrics (unique dates per GSBH per week)
        weekly_metrics = {}
        for (mags, week), group in df_m.groupby(['Mã NV Đăng Ký', 'Tuần']):
            reg_days = group[group['Is_Registered'] == 1]['Từ Ngày'].nunique()
            comp_days = group[group['Is_Completed'] == 1]['Từ Ngày'].nunique()
            weekly_metrics[(mags, week)] = (reg_days, comp_days)
            
        # Monthly total calculation (unique dates across the whole month)
        monthly_metrics = {}
        for mags, group in df_m.groupby('Mã NV Đăng Ký'):
            reg_days = group[group['Is_Registered'] == 1]['Từ Ngày'].nunique()
            comp_days = group[group['Is_Completed'] == 1]['Từ Ngày'].nunique()
            monthly_metrics[mags] = (reg_days, comp_days)
            
        # Determine target based on max date of this month
        max_date = df_m['Từ Ngày'].max() if len(df_m) > 0 else datetime.date.today()
        current_month_target = get_target_day_count_by_date(max_date)
        
        # 7. Build final details table (for detail tab)
        rows_detail = []
        for idx, gs in gsbh_list.iterrows():
            mags = gs['Mã NV Đăng Ký']
            tengs = gs['Tên NV Đăng Ký']
            vung = gs['Tên Vùng']
            mien = gs['Miền']
            
            m_reg, m_comp = monthly_metrics.get(mags, (0, 0))
            m_pct_reg = m_comp / m_reg if m_reg > 0 else 0.0
            
            # Cumulative week targets: W1=4, W2=8, W3=12, W4=16
            weekly_details = {}
            cum_reg = 0
            cum_comp = 0
            for i, wk in enumerate(weeks):
                w_reg, w_comp = weekly_metrics.get((mags, wk), (0, 0))
                
                # Accumulate values
                cum_reg += w_reg
                cum_comp += w_comp
                
                wk_target = (i + 1) * 4.0
                w_pct_reg = cum_comp / cum_reg if cum_reg > 0 else 0.0
                w_pct_target = cum_comp / wk_target
                kpi_status = "Đạt" if cum_comp >= wk_target else "Không đạt"
                
                wk_key = wk.lower().replace(" ", "").replace("tuần", "tuan")
                weekly_details[f"{wk_key}_dang_ky"] = int(w_reg)
                weekly_details[f"{wk_key}_hoan_thanh"] = int(w_comp)
                weekly_details[f"{wk_key}_pct_dang_ky"] = float(w_pct_reg)
                weekly_details[f"{wk_key}_pct_target"] = float(w_pct_target)
                weekly_details[f"{wk_key}_kpi"] = kpi_status
                
            m_pct_target = m_comp / current_month_target if current_month_target > 0 else 0.0
            
            detail_item = {
                "mien": mien,
                "vung": vung,
                "ma_gsbh": mags,
                "ten_gsbh": tengs,
                "dang_ky_thang": int(m_reg),
                "hoan_thanh_thang": int(m_comp),
                "pct_dang_ky": float(m_pct_reg),
                "pct_target": float(m_pct_target)
            }
            detail_item.update(weekly_details)
            rows_detail.append(detail_item)
            
        # 8. Summary calculation
        def get_summary_stats_local(label_mien, label_vung, subset_df, is_mien_total=False, is_grand_total=False):
            total_gs = len(subset_df)
            sum_reg = int(subset_df['dang_ky_thang'].sum())
            sum_comp = int(subset_df['hoan_thanh_thang'].sum())
            pct_reg = sum_comp / sum_reg if sum_reg > 0 else 0.0
            pct_target = sum_comp / (total_gs * current_month_target) if (total_gs > 0 and current_month_target > 0) else 0.0
            
            summary_item = {
                "mien": label_mien,
                "vung": label_vung,
                "is_mien_total": is_mien_total,
                "is_grand_total": is_grand_total,
                "so_gsbh": total_gs,
                "dang_ky": sum_reg,
                "hoan_thanh": sum_comp,
                "pct_dang_ky": float(pct_reg),
                "pct_target": float(pct_target)
            }
            
            for wk in weeks:
                wk_key = wk.lower().replace(" ", "").replace("tuần", "tuan")
                mags_subset = subset_df['ma_gsbh'].tolist()
                num_dat = 0
                for item in rows_detail:
                    if item['ma_gsbh'] in mags_subset and item[f"{wk_key}_kpi"] == "Đạt":
                        num_dat += 1
                pct_dat = num_dat / total_gs if total_gs > 0 else 0.0
                summary_item[f"{wk_key}_dat"] = num_dat
                summary_item[f"{wk_key}_pct"] = float(pct_dat)
                
            return summary_item
            
        summary_list = []
        df_detail_temp = pd.DataFrame(rows_detail)
        
        # Miền Bắc
        df_north = df_detail_temp[df_detail_temp['mien'] == 'Miền Bắc']
        for vung_name, group in df_north.groupby('vung'):
            summary_list.append(get_summary_stats_local('Miền Bắc', vung_name, group))
        summary_list.append(get_summary_stats_local('Miền Bắc', 'TỔNG MIỀN BẮC', df_north, is_mien_total=True))
        
        # Miền Nam
        df_south = df_detail_temp[df_detail_temp['mien'] == 'Miền Nam']
        for vung_name, group in df_south.groupby('vung'):
            summary_list.append(get_summary_stats_local('Miền Nam', vung_name, group))
        summary_list.append(get_summary_stats_local('Miền Nam', 'TỔNG MIỀN NAM', df_south, is_mien_total=True))
        
        # Tổng cộng toàn quốc
        summary_list.append(get_summary_stats_local('TOÀN QUỐC', 'TỔNG CỘNG TOÀN QUỐC', df_detail_temp, is_grand_total=True))
        
        # Save Month Data
        months_data[cur_month] = {
            "currentTargetDays": current_month_target,
            "summary": summary_list,
            "detail": rows_detail
        }
    
    # 9. Extract Filtered Data for auditing
    print("Đang chuẩn bị dữ liệu thô...")
    raw_cols = ['Miền', 'Tên Vùng', 'Mã NV Đăng Ký', 'Tên NV Đăng Ký', 'Từ Ngày', 'Loại Công Việc', 'Trạng Thái Công Việc', 'Tuần']
    df_raw_copy = df_hl_valid[raw_cols].copy()
    df_raw_copy['Từ Ngày'] = df_raw_copy['Từ Ngày'].apply(lambda x: x.strftime('%Y-%m-%d') if hasattr(x, 'strftime') else str(x))
    
    raw_list = []
    for idx, row in df_raw_copy.iterrows():
        raw_list.append({
            "mien": row['Miền'],
            "vung": row['Tên Vùng'],
            "ma_gsbh": row['Mã NV Đăng Ký'],
            "ten_gsbh": row['Tên NV Đăng Ký'],
            "tu_ngay": row['Từ Ngày'],
            "loai_cong_viec": row['Loại Công Việc'],
            "trang_thai": row['Trạng Thái Công Việc'],
            "tuan": row['Tuần']
        })
        
    # 10. Write JavaScript data file
    print("Đang ghi dữ liệu ra file coaching_data.js...")
    now_str = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    
    data_to_export = {
        "lastUpdated": now_str,
        "months": months_data,
        "rawData": raw_list
    }
    
    with open(js_abs, "w", encoding="utf-8") as f:
        f.write("// FILE TỰ ĐỘNG SINH - KHÔNG SỬA TRỰC TIẾP\n")
        f.write(f"const coachingData = {json.dumps(data_to_export, ensure_ascii=False, indent=2)};\n")
        
    print(f"Đã xuất dữ liệu JS thành công tại: {js_output}")
    
    # --- 11. Ghi đè file Excel (để lưu trữ/backup dự phòng) ---
    # Ta vẫn chạy ghi Excel theo format cũ để backup nếu người dùng vẫn cần file Excel
    try:
        shutil.copy2(xlsx_abs, r"d:\CJ\11. Adhoc\T8\WW\Bao_Cao_Coaching_GSBH_backup.xlsx")
    except:
        pass
        
    print("Đang tạo backup Excel dự phòng...")
    latest_month = available_months[-1] if len(available_months) > 0 else None
    if latest_month:
        latest_data = months_data[latest_month]
        with pd.ExcelWriter(xlsx_abs, engine='xlsxwriter') as writer:
            df_summary_excel = pd.DataFrame(latest_data["summary"])
            df_detail_excel = pd.DataFrame(latest_data["detail"])
            df_summary_excel.to_excel(writer, sheet_name='Summary', index=False)
            df_detail_excel.to_excel(writer, sheet_name='Supervisor_Report', index=False)
            df_raw_copy.to_excel(writer, sheet_name='Filtered_Data', index=False)
        
    print("Mọi công đoạn xử lý dữ liệu hoàn tất!")

except Exception as e:
    import traceback
    print("Đã xảy ra lỗi trong quá trình xử lý:")
    traceback.print_exc()
