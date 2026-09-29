import sys
import os
import json
import pandas as pd
from datetime import datetime
import io
import re

import base64
import os
import time
# [BỔ SUNG THƯ VIỆN WEB PUSH]
from pywebpush import webpush, WebPushException

# Cấu hình Khóa VAPID (Thay bằng chuỗi thầy/cô vừa copy ở Bước 1)
VAPID_PUBLIC_KEY = "BGL6lcaXMAK5VhqdjWBO9IZ0ECiKYu47oE_RQE7rCpRinW2Je0IyrnNQi35glDiMPAMzIh0ynLUHxu4aZFCIWSs"
VAPID_PRIVATE_KEY = "eLemgBn1ETfOsv0iT8bA4DfgFH-6poMgySRXaghXohI"
VAPID_CLAIMS = {
    "sub": "mailto:nguyenhoanganh.computer@gmail.com" # Email quản trị của trường
}

# (Tùy chọn) Thầy/cô có thể tạo thêm bảng `PushSubscription` trong Models 
# để lưu thông tin đăng ký của từng GVCN. Tạm thời chúng ta dùng Session/Dict để test.
global_subscriptions = {}
  
def process_and_save_evidence(base64_string, branch_id, week_name):
    """Hàm hứng chuỗi Base64 (Data URI) và đẩy lên Cloudinary một cách an toàn (Đã bọc thép chống nhân đôi)"""
    if not base64_string or base64_string.strip() in ["", "[]"]:
        return None
        
    try:
        import json
        import unicodedata
        import os
        import hashlib
        import cloudinary
        import cloudinary.uploader
        
        # Tự động nhận diện cấu hình Cloudinary
        if os.environ.get("CLOUDINARY_URL"):
            cloudinary.config(secure=True)
        else:
            cloudinary.config( 
                cloud_name = os.environ.get("CLOUDINARY_CLOUD_NAME"), 
                api_key = os.environ.get("CLOUDINARY_API_KEY"), 
                api_secret = os.environ.get("CLOUDINARY_API_SECRET"),
                secure = True
            )
            
        # Xử lý danh sách hoặc chuỗi đơn
        if base64_string.startswith('['):
            base64_list = json.loads(base64_string)
        else:
            base64_list = [base64_string]
            
        saved_paths = []
        
        # Chuẩn hóa tên tuần không dấu
        safe_week = unicodedata.normalize('NFKD', str(week_name)).encode('ASCII', 'ignore').decode('utf-8')
        safe_week = re.sub(r'[^a-zA-Z0-9]', '_', safe_week)
        
        for b64 in base64_list:
            if not b64: continue
            
            # [CHÌA KHÓA VÀNG]: Băm mã Base64 thành Dấu vân tay (MD5) để chống trùng lặp do mạng yếu
            img_hash = hashlib.md5(b64.encode('utf-8')).hexdigest()[:15]
            
            # Đảm bảo chuỗi base64 giữ nguyên định dạng Data URI chuẩn từ FileReader
            # Gắn thêm public_id để Cloudinary tự động ghi đè nếu mạng tự động gửi lại cùng 1 bức ảnh
            upload_result = cloudinary.uploader.upload(
                b64, 
                folder=f"thidua_doantruong/{safe_week}",
                public_id=f"img_{branch_id}_{img_hash}"
            )
            saved_paths.append(upload_result['secure_url'])
            
        return "|".join(saved_paths) if saved_paths else None
        
    except Exception as e:
        error_detail = str(e)
        print(f"❌ LỖI UPLOAD ẢNH LÊN ĐÁM MÂY: {error_detail}")
        raise Exception(error_detail)
    
from flask import Flask, render_template, request, redirect, url_for, flash, session, send_file
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.worksheet.page import PageMargins
from openpyxl.utils import get_column_letter

# Import các thành phần từ thư mục con của dự án
from database.database import init_db, get_session, session_scope
from database.models import User, UserRole, SchoolYear, Branch, RedStar, DutyArea, Assignment, StarEvaluation, WeeklyScore, ViolationCategory, WeeklyViolation, RawScore, MonthlyRecord, ActionLog, ScoreSettings, GVCNAttendance

# Import hàm kiểm tra đăng nhập từ file account_manager
from database.account_manager import verify_external_login, sync_account_to_json, remove_account_from_json, load_external_accounts, save_external_accounts

# Khởi tạo ứng dụng Web Flask
app = Flask(__name__)
# Thiết lập khóa bí mật để sử dụng Flash messages báo lỗi
app.secret_key = "doan_truong_thanh_hoa_secret_key" 
# [BỔ SUNG]: API LƯU KHÓA ĐĂNG KÝ VÀ HÀM PHÁT SÓNG
@app.route('/api/save_subscription', methods=['POST'])
def save_subscription():
    """Lưu mã đăng ký nhận thông báo của thiết bị"""
    sub_info = request.get_json()
    username = session.get('username')
    
    if not username or not sub_info:
        return {"success": False, "error": "Lỗi dữ liệu"}, 400
        
    global_subscriptions[username] = sub_info
    return {"success": True, "message": "Đã kết nối luồng thông báo đẩy!"}

def send_web_push(username, title, body):
    """Hàm lõi gọi để bắn thông báo tới 1 user cụ thể"""
    sub_info = global_subscriptions.get(username)
    if not sub_info:
        return False
        
    try:
        webpush(
            subscription_info=sub_info,
            data=json.dumps({"title": title, "body": body}),
            vapid_private_key=VAPID_PRIVATE_KEY,
            vapid_claims=VAPID_CLAIMS
        )
        return True
    except WebPushException as ex:
        print("Lỗi gửi Push:", repr(ex))
        return False
    
from flask import g

# ==========================================
# --- HÀM HỖ TRỢ: GHI NHẬT KÝ THAO TÁC HỆ THỐNG ---
# ==========================================
def log_system_action(action_type, details=""):
    """
    Hàm lưu vết thao tác vào bộ nhớ đệm (Flask g). 
    Sẽ được ghi vào database SAU KHI xử lý xong luồng chính để tránh lỗi "database is locked".
    """
    if 'action_logs' not in g:
        g.action_logs = []
        
    g.action_logs.append({
        'username': session.get('username', 'Khách'),
        'full_name': session.get('full_name', 'Người dùng Ẩn danh'),
        'action_type': action_type,
        'details': details,
        'timestamp': datetime.utcnow()
    })

@app.after_request
def save_action_logs(response):
    """
    Tự động kích hoạt sau mỗi request: Đổ dữ liệu log từ bộ nhớ đệm vào Database.
    Vì lúc này Database đã được luồng chính giải phóng, nên sẽ KHÔNG BAO GIỜ bị lỗi Locked.
    """
    if hasattr(g, 'action_logs') and g.action_logs:
        try:
            from database.database import session_scope
            from database.models import ActionLog
            
            with session_scope() as db_session:
                for log_data in g.action_logs:
                    new_log = ActionLog(
                        username=log_data['username'],
                        full_name=log_data['full_name'],
                        action_type=log_data['action_type'],
                        details=log_data['details'],
                        timestamp=log_data['timestamp']
                    )
                    db_session.add(new_log)
                db_session.commit()
        except Exception as e:
            print(f"Lỗi lưu file log: {e}")
    return response

# ==========================================
# CÁC HÀM TỰ ĐỘNG KHỞI TẠO VÀ ĐỒNG BỘ HỆ THỐNG
# ==========================================
def auto_init_accounts():
    """Hàm tự động quét CSDL và đồng bộ lại file accounts.json để chống lỗi đăng nhập trên Render"""
    try:
        from database.database import session_scope
        from database.models import User
        import os, json
        
        data_folder = "data"
        os.makedirs(data_folder, exist_ok=True)
        file_path = os.path.join(data_folder, "accounts.json")
        
        with session_scope() as db_session:
            # 1. Lấy toàn bộ user đang có trong CSDL
            users = db_session.query(User).all()
            
            # Nếu CSDL hoàn toàn trống, tạo 1 acc admin mặc định
            if not users:
                default_accounts = [{
                    "username": "admin", "password": "1",  
                    "full_name": "Bí thư Đoàn trường", "role": "Bí thư"
                }]
                with open(file_path, "w", encoding="utf-8") as f:
                    json.dump(default_accounts, f, indent=4, ensure_ascii=False)
                print("Hệ thống: Đã khởi tạo accounts.json mặc định.")
                return

            # 2. Nếu CSDL có dữ liệu, ép đồng bộ toàn bộ từ CSDL -> file JSON
            accounts_data = []
            for u in users:
                role_str = u.role.value if hasattr(u.role, 'value') else str(u.role)
                accounts_data.append({
                    "username": u.username,
                    "password": u.password_hash,
                    "full_name": u.full_name,
                    "role": role_str,
                    "is_active": getattr(u, 'is_active', True)
                })
                
            with open(file_path, "w", encoding="utf-8") as f:
                json.dump(accounts_data, f, indent=4, ensure_ascii=False)
                
            print(f"✅ HỆ THỐNG: Đã đồng bộ {len(users)} tài khoản từ Database sang JSON thành công!")
            
    except Exception as e:
        print(f"❌ Lỗi đồng bộ tài khoản: {e}")

# KÍCH HOẠT NGAY LẬP TỨC KHI RENDER CHẠY
auto_init_accounts()

def create_mock_admin():
    """Tạo một tài khoản Bí thư Đoàn trường mặc định trong SQLite để test"""
    try:
        with session_scope() as db_session:
            admin = db_session.query(User).filter(User.username == "admin").first()
            if not admin:
                new_admin = User(
                    username="admin",
                    password_hash="1",  
                    full_name="Bí thư Đoàn trường",
                    role=UserRole.BI_THU,
                    is_active=True
                )
                db_session.add(new_admin)
                print("Đã khởi tạo tài khoản thành công: User: admin | Pass: 1")
    except Exception as e:
        print(f"Lỗi khởi tạo tài khoản admin: {e}")

@app.before_request
def restrict_access():
    # 1. [BẢN VÁ LỖI PWA]: SỬ DỤNG .path THAY VÌ .endpoint ĐỂ CHỐNG REDIRECT CHO APP MOBILE
    if request.path in ['/sw.js', '/manifest.json'] or request.path.startswith('/static/'):
        return None
        
    # 2. KHAI BÁO DANH SÁCH QUYỀN TRUY CẬP CHO TỪNG ROLE
    allowed_for_gvcn = [
        'login', 'logout', 'ping', 'change_password', 
        'class_dashboard', 'preview_class_dashboard', 'export_class_dashboard', 'submit_appeal',
        'parse_sodaubai', 
        'api_submit_evaluation',
        'api_gvcn_leaderboard',
        'api_gvcn_get_months',
        'update_branch_info',
        'api_weekly_scores_json',
        'gvcn_checkin',
        'api_class_blacklist',      # <--- BỔ SUNG DÒNG NÀY
        'export_class_blacklist'    # <--- BỔ SUNG DÒNG NÀY
    ]
    
    allowed_for_saodo = [
        'login', 'logout', 'ping', 'change_password', 
        'mobile_sao_do', '', 
        'quick_log_violation_form', 'handle_raw_scores',
        'sao_do_quick_submit_form', 'submit_mobile_sao_do',
        'api_submit_evaluation',
        'export_filtered_blacklist', # [BỔ SUNG]: Cho phép Sao đỏ xuất file bản lọc nếu cần
        'api_weekly_scores_json'
    ]

    allowed_for_bgh = [
        'login', 'logout', 'ping', 'change_password', 
        'bgh_dashboard',                 
        'class_dashboard',               
        'preview_class_dashboard',       
        'export_class_dashboard',        
        'class_monthly_analysis',        
        'class_semester_analysis',       
        'school_monthly_analysis',
        'gvcn_attendance_stats'
    ]
    
    # 3. KIỂM TRA QUYỀN TRUY CẬP THEO ROLE
    role = session.get('role')
    
    # Kiểm tra quyền Giáo viên chủ nhiệm
    if role == 'Giáo viên chủ nhiệm':
        if request.endpoint and request.endpoint not in allowed_for_gvcn:
            if request.path.startswith('/api/'):
                return {
                    "success": False,
                    "error": "Bạn không có quyền truy cập API này.",
                    "endpoint": request.endpoint
                }, 403

            flash("⛔ Từ chối truy cập: Quyền GVCN!", "error")
            return redirect(url_for('class_dashboard'))

    # [BỔ SUNG QUAN TRỌNG]: Kiểm tra quyền Đội Sao Đỏ (Chặn không cho đi lạc trang khác)
    elif role == 'Sao đỏ':
        if request.endpoint and request.endpoint not in allowed_for_saodo:
            if request.path.startswith('/api/'):
                return {
                    "success": False,
                    "error": "Sao đỏ không có quyền truy cập API này.",
                    "endpoint": request.endpoint
                }, 403

            flash("⛔ Từ chối truy cập: Bạn chỉ được phép sử dụng App trực cổng!", "error")
            return redirect(url_for('mobile_sao_do'))

    # Kiểm tra quyền Ban Giám hiệu (Chỉ cho phép xem báo cáo và điều hành)
    elif role == 'Ban Giám hiệu':
        if request.endpoint and request.endpoint not in allowed_for_bgh:
            if request.path.startswith('/api/'):
                return {
                    "success": False,
                    "error": "Ban Giám hiệu chỉ có quyền xem báo cáo và điều hành.",
                    "endpoint": request.endpoint
                }, 403

            flash("⛔ Từ chối truy cập: Tài khoản BGH chỉ có quyền xem báo cáo và điều hành!", "error")
            return redirect(url_for('bgh_dashboard'))

@app.route('/ping')
def ping():
    return "<h1>Kết nối thành công! Máy chủ đang hoạt động.</h1>"

@app.route('/api/gvcn_checkin', methods=['POST'])
def gvcn_checkin():
    # Chỉ GVCN mới được phép điểm danh
    if session.get('role') != 'Giáo viên chủ nhiệm':
        return {"success": False, "error": "Không có quyền thực hiện!"}, 403

    data = request.get_json(force=True, silent=True) or request.form.to_dict()
    branch_id = data.get('branch_id')
    week_name = data.get('week_name')

    if not branch_id:
        return {"success": False, "error": "Thiếu thông tin lớp!"}, 400

    try:
        from datetime import datetime, timezone, timedelta
        vn_tz = timezone(timedelta(hours=7))
        
        with session_scope() as db_session:
            today_date = datetime.now(vn_tz).date()
            
            # =========================================================================
            # [BẢN VÁ LỖI CỐT LÕI]: Chặn đứng chuỗi "Tuần hiện tại" từ Mobile App gửi lên
            # Tự động đồng bộ tên tuần với lịch trực / điểm thi đua
            # =========================================================================
            if not week_name or week_name == "Tuần hiện tại" or "Tuần" not in str(week_name):
                latest_assign = db_session.query(Assignment).order_by(Assignment.week_number.desc()).first()
                if latest_assign:
                    week_name = f"Tuần {latest_assign.week_number}"
                else:
                    latest_score = db_session.query(WeeklyScore).order_by(WeeklyScore.id.desc()).first()
                    if latest_score:
                        week_name = latest_score.week
                    else:
                        week_name = "Tuần 1"
            
            # Kiểm tra xem hôm nay thầy cô đã bấm chưa
            exist = db_session.query(GVCNAttendance).filter_by(branch_id=branch_id, date=today_date).first()
            if exist:
                return {"success": False, "error": "Thầy/Cô đã điểm danh cho ngày hôm nay rồi!"}, 400

            # Ghi nhận vào DB
            new_attendance = GVCNAttendance(
                branch_id=branch_id, 
                week_name=week_name, 
                date=today_date
            )
            db_session.add(new_attendance)
            
            log_system_action("ĐIỂM DANH GVCN", f"GVCN Lớp ID {branch_id} đã điểm danh sinh hoạt 15p {week_name}")
            
            return {"success": True, "message": "✅ Điểm danh thành công! Cảm ơn Thầy/Cô."}, 200
    except Exception as e:
        import traceback; traceback.print_exc()
        return {"success": False, "error": str(e)}, 500
    
@app.route('/gvcn_attendance_stats')
def gvcn_attendance_stats():
    # Chỉ Admin, BGH hoặc Bí thư mới được xem thống kê này
    if session.get('role') not in ['Quản trị viên', 'Admin', 'Bí thư Đoàn trường', 'Bí thư', 'Ban Giám hiệu']:
        flash("Bạn không có quyền xem bảng thống kê này!", "error")
        return redirect(url_for('dashboard'))
        
    try:
        with session_scope() as db_session:
            active_year = db_session.query(SchoolYear).filter_by(is_active=True).first()
            if not active_year:
                flash("Chưa có năm học kích hoạt!", "error")
                return redirect(url_for('dashboard'))
                
            time_mode = request.args.get('time_mode', 'week') 
            time_value = request.args.get('time_value', '')
            
            # Lấy toàn bộ dữ liệu điểm danh của năm học hiện tại
            all_atts = db_session.query(GVCNAttendance).join(Branch).filter(Branch.school_year_id == active_year.id).all()
            
            # =====================================================================
            # [THUẬT TOÁN CHỮA BỆNH]: TỰ ĐỘNG QUÉT VÀ SỬA LỖI TÊN TUẦN TRONG CSDL
            # Giúp bóc tách T2(14/09) và T2(21/09) về đúng 2 Tuần riêng biệt
            # =====================================================================
            import datetime as dt
            changes_made = False
            known_weeks = {}
            
            # Lấy mốc chuẩn từ Lịch phân công trực
            asm_records = db_session.query(Assignment).filter(Assignment.date != None).all()
            for asm in asm_records:
                try:
                    py_date = asm.date
                    if isinstance(py_date, str):
                        py_date = dt.datetime.strptime(py_date.split()[0], '%Y-%m-%d').date()
                    monday = py_date - dt.timedelta(days=py_date.weekday())
                    known_weeks[monday] = f"Tuần {asm.week_number}"
                except: pass

            base_monday = min(known_weeks.keys()) if known_weeks else None
            if not base_monday and all_atts:
                valid_dates = [a.date for a in all_atts if a.date]
                if valid_dates:
                    base_monday = min(valid_dates) - dt.timedelta(days=min(valid_dates).weekday())

            # Chữa lỗi cho các bản ghi bị gắn mác "Tuần hiện tại"
            for a in all_atts:
                if (not a.week_name or a.week_name == "Tuần hiện tại" or "Tuần" not in str(a.week_name)) and a.date:
                    a_monday = a.date - dt.timedelta(days=a.date.weekday())
                    if a_monday in known_weeks:
                        a.week_name = known_weeks[a_monday]
                    elif base_monday:
                        week_num = ((a_monday - base_monday).days // 7) + 1
                        a.week_name = f"Tuần {week_num}"
                    else:
                        a.week_name = f"Tuần {a.date.isocalendar()[1] - 34}"
                    changes_made = True
            
            if changes_made:
                db_session.commit()
            # =====================================================================
            
            # Tự động trích xuất các Tuần, Tháng, Học kỳ, Năm học đã có dữ liệu để làm bộ lọc
            import re
            available_weeks = sorted(list(set([a.week_name for a in all_atts if a.week_name])), key=lambda x: int(''.join(filter(str.isdigit, x))) if any(c.isdigit() for c in x) else 0)
            available_months = sorted(list(set([a.date.strftime('Tháng %m/%Y') for a in all_atts if a.date])), reverse=True)
            
            available_semesters = set()
            available_years = set()
            
            for a in all_atts:
                if a.date:
                    start_year = a.date.year if a.date.month >= 8 else a.date.year - 1
                    school_year_str = f"{start_year}-{start_year + 1}"
                    
                    hk_str = f"Học kỳ 1 ({school_year_str})" if a.date.month >= 8 or a.date.month == 1 else f"Học kỳ 2 ({school_year_str})"
                    available_semesters.add(hk_str)
                    available_years.add(f"Năm học {school_year_str}")

            available_semesters = sorted(list(available_semesters), reverse=True)
            available_years = sorted(list(available_years), reverse=True)

            # Đặt giá trị mặc định khi vừa vào trang -> Ưu tiên TUẦN MỚI NHẤT
            if time_mode == 'week' and not time_value and available_weeks:
                time_value = available_weeks[-1] 
            elif time_mode == 'month' and not time_value and available_months:
                time_value = available_months[0]
            elif time_mode == 'semester' and not time_value and available_semesters:
                time_value = available_semesters[0]
            elif time_mode == 'year' and not time_value and available_years:
                time_value = available_years[0]
                
            # Bộ lọc dữ liệu (Chỉ lấy đúng mốc BGH chọn)
            filtered_atts = []
            for a in all_atts:
                if not a.date: continue
                if time_mode == 'week' and a.week_name == time_value:
                    filtered_atts.append(a)
                elif time_mode == 'month' and a.date.strftime('Tháng %m/%Y') == time_value:
                    filtered_atts.append(a)
                elif time_mode == 'semester':
                    start_year = a.date.year if a.date.month >= 8 else a.date.year - 1
                    hk_str = f"Học kỳ 1 ({start_year}-{start_year + 1})" if a.date.month >= 8 or a.date.month == 1 else f"Học kỳ 2 ({start_year}-{start_year + 1})"
                    if hk_str == time_value:
                        filtered_atts.append(a)
                elif time_mode == 'year':
                    start_year = a.date.year if a.date.month >= 8 else a.date.year - 1
                    if f"Năm học {start_year}-{start_year + 1}" == time_value:
                        filtered_atts.append(a)
                    
            # NHÓM DỮ LIỆU ĐIỂM DANH THEO TỪNG TUẦN THI ĐUA
            stats = {}
            branches = db_session.query(Branch).filter_by(school_year_id=active_year.id).all()
            for b in branches:
                stats[b.id] = {
                    'branch_name': b.name,
                    'gvcn': b.gvcn or "Chưa cập nhật",
                    'count': 0,
                    'weeks': {} # Dictionary chứa số buổi tách biệt từng tuần
                }
            
            day_map = {0: 'T2', 1: 'T3', 2: 'T4', 3: 'T5', 4: 'T6', 5: 'T7', 6: 'CN'}   
            
            for a in filtered_atts:
                if a.branch_id in stats and a.date:
                    stats[a.branch_id]['count'] += 1 
                    
                    day_str = day_map.get(a.date.weekday(), '')
                    date_str = f"{day_str} ({a.date.strftime('%d/%m')})" 
                    
                    week_key = a.week_name or "Khác"
                    if week_key not in stats[a.branch_id]['weeks']:
                        stats[a.branch_id]['weeks'][week_key] = []
                        
                    stats[a.branch_id]['weeks'][week_key].append(date_str)                    
            
            # Sắp xếp các tuần bên trong từng lớp cho chuẩn (VD: Tuần 1 hiển thị trước Tuần 2)
            for b_id, b_data in stats.items():
                sorted_weeks = {}
                for w in sorted(b_data['weeks'].keys(), key=lambda x: int(''.join(filter(str.isdigit, x))) if any(c.isdigit() for c in x) else 0):
                    sorted_weeks[w] = b_data['weeks'][w]
                b_data['weeks'] = sorted_weeks

            # Sắp xếp tự nhiên tên lớp 10A2 đứng trước 10A10
            stats_list = list(stats.values())
            stats_list.sort(key=lambda x: [int(t) if t.isdigit() else t.lower() for t in re.split(r'(\d+)', str(x['branch_name']))])
            
            return render_template('gvcn_attendance.html', 
                                   stats_list=stats_list,
                                   available_weeks=available_weeks,
                                   available_months=available_months,
                                   available_semesters=available_semesters,
                                   available_years=available_years,
                                   time_mode=time_mode,
                                   time_value=time_value)
    except Exception as e:
        import traceback; traceback.print_exc()
        flash(f"Lỗi tải thống kê: {e}", "error")
        return redirect(url_for('dashboard'))
    
# API: BÓC TÁCH DỮ LIỆU TỪ FILE SỔ ĐẦU BÀI (TỐI ƯU CHỐNG SÓT ĐIỂM VÀ LỌC THEO NGÀY)
@app.route('/api/parse_sodaubai', methods=['POST'])
@app.route('/weekly/api/parse_sodaubai', methods=['POST'])
def parse_sodaubai():
    if 'excel_file' not in request.files:
        return {"error": "Không tìm thấy file!"}, 400
        
    file = request.files['excel_file']
    if file.filename == '':
        return {"error": "Chưa chọn file nào!"}, 400
        
    # Lấy tên lớp dự kiến từ giao diện để đối chiếu
    expected_branch = request.form.get('expected_branch_name', '').strip().upper()
    
    # [NÂNG CẤP LÕI 1]: Nhận tham số ngày cần quét từ giao diện (Mặc định "Tất cả")
    target_day_input = request.form.get('target_day', 'Tất cả').strip()
        
    try:
        import pandas as pd
        
        
        df = pd.read_excel(file, header=None)
        
        # =================================================================
        # THUẬT TOÁN KHIÊN BẢO VỆ & NHẬN DIỆN LỚP TỰ ĐỘNG
        # =================================================================
        found_class_name = None
        # Quét tối đa 50 dòng đầu và toàn bộ cột để tìm chữ "Lớp: ..."
        for i in range(min(50, len(df))):
            for j in range(len(df.columns)):
                cell_val = str(df.iloc[i, j]).strip()
                if cell_val.lower().startswith('lớp:'):
                    match = re.search(r'Lớp:\s*([A-Za-z0-9]+)', cell_val, re.IGNORECASE)
                    if match:
                        found_class_name = match.group(1).upper()
                        break
            if found_class_name:
                break
        
        # Nếu đang quét đơn lẻ (có expected_branch) thì khóa nòng kiểm tra
        if expected_branch and found_class_name and found_class_name != expected_branch:
            return {
                "error": f"⛔ CẢNH BÁO: FILE SỔ ĐẦU BÀI KHÔNG KHỚP!\nBạn đang ở form nhập điểm của lớp {expected_branch}, nhưng file Excel bạn vừa tải lên lại là Sổ đầu bài của lớp {found_class_name}. Vui lòng chọn lại đúng file!"
            }, 400
        
        # Nếu quét hàng loạt nhưng không tìm thấy tên lớp trong file
        if not expected_branch and not found_class_name:
            return {"error": "Không nhận diện được Tên lớp trong file Excel này (Thiếu ô 'Lớp: ...')."}, 400
        # =================================================================
        
        try:
            start_row = df[df[0].astype(str).str.contains('Thứ \nngày tháng', na=False, case=False)].index[0]
        except:
            return {"error": "Hệ thống không nhận diện được biểu mẫu Sổ Đầu Bài này!"}, 400
        # =========================================================
        # [BẢN VÁ LỖI]: TỰ ĐỘNG DÒ TÌM CỘT ĐIỂM VÀ NHẬN XÉT
        # =========================================================
        col_diem_list = []
        col_xep_loai_tiet = 18 # Cột mặc định dự phòng
        
        # Quét ngang các cột ở dòng tiêu đề (start_row) để tìm đúng vị trí
        for c in range(len(df.columns)):
            col_title = str(df.iloc[start_row, c]).lower()
            
            # [ĐÃ VÁ LỖI]: Không lấy cột nếu tiêu đề có chứa chữ "xếp loại"
            if ("điểm" in col_title or "nhận xét" in col_title) and "xếp loại" not in col_title:
                col_diem_list.append(c)
            
            # Dò tìm cột Xếp loại độc lập
            if "xếp loại" in col_title:
                col_xep_loai_tiet = c
                
        # Nếu biểu mẫu quá lạ không dò ra chữ, quay về giá trị mặc định
        if not col_diem_list: 
            col_diem_list = [13, 14, 15]
            
        c10 = c9 = c8 = 0
        
        subject_scores = {}  # Phân loại điểm tốt (8, 9, 10) theo Tên Môn Học
        bad_marks_list = []  # Lưu tạm toàn bộ lỗi điểm kém/không học bài để xén trần
        
        # [BẢN VÁ LỖI]: Dùng SET để lọc trùng lặp học sinh vắng trong cùng 1 ngày
        general_violations_set = set()
        current_day = "Ngày khác"
        
        cat_khb = "Không học bài"
        cat_dk = "Bị điểm kém"
        violation_names = [] # Khởi tạo danh sách tên lỗi an toàn
        try:
            from database.models import ViolationCategory, SchoolYear
            from database.database import session_scope
            with session_scope() as db_session:
                active_year = db_session.query(SchoolYear).filter_by(is_active=True).first()
                cats = db_session.query(ViolationCategory).filter_by(school_year_id=active_year.id).all() if active_year else []
                
                # Sắp xếp tên từ dài đến ngắn để AI không nhận diện nhầm lỗi con
                violation_names = sorted([c.name for c in cats], key=len, reverse=True) 
                
                for c in cats:
                    nl = c.name.lower()
                    if "không học" in nl or "không thuộc" in nl: cat_khb = c.name
                    if "điểm kém" in nl or "điểm yếu" in nl or "điểm 0" in nl: cat_dk = c.name
        except:
            pass
        # ---------------------------------------------------------------------------------
        # Bắt đầu quét từ start_row + 3 (bỏ qua dòng tiêu đề và hàng số thứ tự 1-10)
        for i in range(start_row + 3, len(df)):
            if i >= len(df): break
            
            cot0_text = str(df.iloc[i, 0])
            if "Ý kiến nhận xét" in cot0_text or "Tổng số tiết" in cot0_text: break

            # --- [THÊM TÍNH NĂNG]: Theo dõi ngày hiện tại để chống lặp ---
            cot0_clean = cot0_text.strip()
            if cot0_clean and cot0_clean.lower() != 'nan':
                current_day = cot0_clean.split('\n')[0].strip()

            # =========================================================================
            # [NÂNG CẤP LÕI 2]: BỘ LỌC CHỈ QUÉT THEO NGÀY CHỈ ĐỊNH (HỖ TRỢ CẢ SÁNG/CHIỀU)
            # =========================================================================
            if target_day_input != 'Tất cả':
                # Chuẩn hóa đầu vào để so sánh linh hoạt (VD: "Thứ 5 Chiều" -> "thứ 5")
                base_target_day = target_day_input.lower().replace(' sáng', '').replace(' chiều', '').strip()
                
                # 1. Nếu dòng trên Sổ đầu bài không có chứa "Thứ 5" -> Lập tức bỏ qua
                if base_target_day not in current_day.lower():
                    continue 
                    
                # 2. Xử lý phân biệt Sáng / Chiều chéo nhau tự động cho mọi ngày
                # Nếu người dùng chọn Sáng nhưng dòng trong sổ là Chiều -> Bỏ qua
                if 'sáng' in target_day_input.lower() and ('chiều' in current_day.lower() or 'chieu' in current_day.lower()):
                    continue
                # Nếu người dùng chọn Chiều nhưng dòng trong sổ là Sáng -> Bỏ qua
                if 'chiều' in target_day_input.lower() and 'sáng' in current_day.lower():
                    continue
            # =========================================================================
            # =============================================================================
            # --- [BỔ SUNG BƯỚC 2]: TỰ ĐỘNG BẮT LỖI VẮNG HỌC (THÔNG MINH HƠN + HỖ TRỢ K/P) ---
            # =============================================================================
            try:
                # Tự động dò tìm cột "Tên HS nghỉ tiết" hoặc "Vắng"
                col_vang = 7 
                for c in range(len(df.columns)):
                    col_title = str(df.iloc[start_row, c]).lower()
                    if "nghỉ tiết" in col_title or "vắng" in col_title:
                        col_vang = c
                        break
                        
                val_vang = str(df.iloc[i, col_vang]).strip()
                if val_vang and val_vang.lower() != 'nan':
                    val_vang = re.sub(r'(?i)vắng\s*(không\s*phép|kp|có\s*phép|cp)?\s*[:\-\(]?\s*0\s*[\)]?\s*(hs|học\s*sinh)?', '', val_vang)
                    # Cắt chuỗi theo dấu phẩy/chấm phẩy để xử lý từng học sinh
                    for p in re.split(r'[,;\n]+', val_vang):
                        p = p.strip()
                        if not p: continue
                        p_lower = p.lower()
                        
                        is_co_phep = False
                        is_khong_phep = False
                        
                        # 1. Nhận diện VẮNG CÓ PHÉP (có chứa chữ p, cp, có phép, ốm, bệnh)
                        if re.search(r'\b(cp|p|có phép|co phep|ốm|bệnh)\b', p_lower):
                            is_co_phep = True
                        
                        # 2. Nhận diện VẮNG KHÔNG PHÉP (có chứa chữ k, kp, không phép, ko phép)
                        elif re.search(r'\b(kp|k|không phép|khong phep|ko phép|ko phep|k phép)\b', p_lower) or 'không' in p_lower:
                            is_khong_phep = True
                            
                        # 3. NẾU GIÁO VIÊN CHỈ GHI TÊN (Không ghi chú gì thêm) -> Mặc định là Không phép
                        else:
                            is_khong_phep = True 
                        
                        # Xóa bỏ các thông tin thừa trong ngoặc và các từ viết tắt để lấy đúng Tên HS (Đã thêm "k" và "p")
                        stu_name = re.sub(r'\(.*?\)', '', p)
                        stu_name = re.sub(r'(?i)\b(không phép|khong phep|ko phép|ko phep|k phép|kp|k|có phép|co phep|cp|p|không|ko|ốm|bệnh)\b', '', stu_name)
                        stu_name = stu_name.strip(' -:').title()
                        
                        if stu_name:
                            if is_khong_phep:
                                general_violations_set.add(('Vắng học không phép', stu_name, current_day))
                            elif is_co_phep:
                                general_violations_set.add(('Vắng học có phép', stu_name, current_day))
            except Exception as e:
                pass
            # =============================================================================
                
            # Lấy tên môn học (Cột 3 theo biểu mẫu Sổ đầu bài)
            try: mon = str(df.iloc[i, 3]).strip()
            except: mon = "Khác"
            if not mon or mon.lower() == 'nan': mon = "Khác"

            # =========================================================================
            # --- [BẢN VÁ LỖI]: TÁCH RIÊNG CỘT XẾP LOẠI (18) KHỎI LUỒNG QUÉT CHỮ ---
            # =========================================================================
            # 1. Quét riêng cột 18 để phạt tập thể "Tiết Yếu" (nếu có)
            try:
                if col_xep_loai_tiet < len(df.columns):
                    xep_loai_tiet = str(df.iloc[i, col_xep_loai_tiet]).strip().lower()
                    if xep_loai_tiet in ['yếu', 'kém']:
                        general_violations_set.add(('Tiết Yếu', '', current_day))
            except Exception:
                pass

            # 2. CHỈ gộp Cột 13, 14 (Điểm KT) và 15 (Nhận xét) để AI dò chữ và bắt lỗi cá nhân
            row_scores = []
            for col in col_diem_list:
                if col < len(df.columns):
                    val = str(df.iloc[i, col]).strip()
                    if val.lower() != 'nan': row_scores.append(val)
                    
            # [BẢN VÁ LỖI TỐI THƯỢNG]: Dùng dấu CHẤM PHẨY để tạo vách ngăn giữa các cột
            diem_raw = " ; ".join(row_scores)
            # =========================================================================
            diem_raw = re.sub(r'(?i)vắng\s*(không\s*phép|kp|có\s*phép|cp)?\s*[:\-\(]?\s*0\s*[\)]?\s*(hs|học\s*sinh)?', '', diem_raw)
            if not diem_raw or 'Ý kiến' in diem_raw or 'BAN GIÁM' in diem_raw:
                continue
                
            # THUẬT TOÁN MỚI: Tách theo dấu phẩy/chấm phẩy, trích xuất điểm bất chấp có nhận xét kèm theo phía sau
            entries = re.split(r'[,;]+', diem_raw)
            parsed_any = False
            
            for entry in entries:
                entry = entry.strip()
                if not entry: continue
                
                # =========================================================================
                # --- [BẢN VÁ TỐI THƯỢNG]: DÒ TÌM LỖI BẰNG CHỮ THEO DANH SÁCH DATABASE ---
                # =========================================================================
                found_text_violation = False
                for v_name in violation_names:
                    v_name_lower = v_name.lower()
                    
                    # Bỏ qua 2 lỗi này vì đã có thuật toán quét bằng "Số điểm" cực mạnh ở dưới
                    if v_name_lower in ['không học bài', 'bị điểm kém']: continue 
                    
                    if v_name_lower in entry.lower():
                        stu_name = ""
                        # Ưu tiên 1: Tìm tên học sinh nằm trong ngoặc vuông hoặc tròn
                        match_bracket = re.search(r'\[(.*?)\]|\((.*?)\)', entry)
                        if match_bracket:
                            stu_name = match_bracket.group(1) or match_bracket.group(2)
                        else:
                            # Ưu tiên 2: Tìm tên học sinh đứng trước/sau dấu phân cách hoặc lấy phần chữ còn lại
                            clean_name = re.sub(v_name, '', entry, flags=re.IGNORECASE)
                            clean_name = re.sub(r'[:\-x0-9]', '', clean_name).strip()
                            if len(clean_name) > 0 and len(clean_name) <= 25: 
                                stu_name = clean_name
                        
                        stu_name = stu_name.strip()
                        
                        # [BẢN VÁ LỖI NÒNG CỐT]: Chẻ nhỏ tên học sinh nếu bị dính chùm
                        if stu_name:
                            
                            split_names = re.split(r'[,;]|\s+và\s+|\s+&\s+|\s{2,}', stu_name, flags=re.IGNORECASE)
                            
                            for s_name in split_names:
                                s_name = s_name.strip().title()
                                if s_name:
                                    general_violations_set.add((v_name, s_name, current_day))
                        else:
                            # Không ghi tên ai thì phạt chung tập thể lớp
                            general_violations_set.add((v_name, "", current_day))
                            
                        found_text_violation = True
                        break # Đã chốt được lỗi cho cụm từ này thì dừng vòng lặp quét chữ
                        
                if found_text_violation:
                    continue # Đã là lỗi bằng chữ thì bỏ qua, không quét điểm số nữa để tránh nhầm lẫn
                # =========================================================================
                # [ĐÃ NÂNG CẤP]: Cho phép thêm [0-9] vào phần tên học sinh
                match = re.search(r'([A-ZÀ-Ỹa-zà-ỹ0-9\s]+?)\s*[:\-]?\s*(\+)?\s*(10|[0-9])\s*(\+)?\s*(?:đ|Đ|điểm|Điểm)?(?!\d)', entry)
                if match:
                    parsed_any = True
                    raw_name = match.group(1).strip()
                    has_plus_before = match.group(2) # Dấu cộng đứng trước (VD: +1)
                    score_val = int(match.group(3))
                    has_plus_after = match.group(4)  # Dấu cộng đứng sau (VD: 1+)
                    
                    # [CHỐNG TRỪ ĐIỂM OAN]: Có dấu + ở trước HOẶC sau đều bỏ qua (không trừ điểm)
                    if (has_plus_before == '+' or has_plus_after == '+') and score_val in [1, 2]:
                        continue     
                    # Tách các từ ra để lọc
                    name_words = raw_name.split()             
                    # [BẢN VÁ LỖI]: Danh sách các từ vô nghĩa cần loại bỏ khi giáo viên ghi nhận xét
                    stop_words = ['không', 'thuộc', 'bài', 'kém', 'lười', 'chú', 'ý', 'phát', 'biểu', 'ồn', 'tập', 'trung', 'nói', 'chuyện', 'đùa', 'giỡn', 'mất', 'trật', 'tự', 'thiếu', 'ngủ', 'quên']                  
                    # Lọc bỏ các từ nằm trong stop_words (không phân biệt hoa/thường)
                    filtered_words = [w for w in name_words if w.lower() not in stop_words]              
                    # Lấy từ cuối cùng trong danh sách ĐÃ LỌC SẠCH làm tên học sinh
                    name_part = filtered_words[-1].title() if filtered_words else "Học sinh"           
                    
                    tiet = str(df.iloc[i, 2]).strip()
                    tiet_str = f"Tiết {tiet}" if tiet != 'nan' else "Tiết học"
                    
                    if score_val == 10:
                        c10 += 1
                        if mon not in subject_scores: subject_scores[mon] = {'c10': 0, 'c9': 0, 'c8': 0}
                        subject_scores[mon]['c10'] += 1
                    elif score_val == 9:
                        c9 += 1
                        if mon not in subject_scores: subject_scores[mon] = {'c10': 0, 'c9': 0, 'c8': 0}
                        subject_scores[mon]['c9'] += 1
                    elif score_val == 8:
                        c8 += 1
                        if mon not in subject_scores: subject_scores[mon] = {'c10': 0, 'c9': 0, 'c8': 0}
                        subject_scores[mon]['c8'] += 1
                        
                    elif score_val == 0:
                        key = f"{name_part} (Môn {mon})" if name_part else f"Môn {mon}"
                        bad_marks_list.append({'type': cat_khb, 'key': key, 'mon': mon}) # Lỗi 0 điểm (Không học bài)
                    elif score_val in [1, 2]: # <--- ĐÃ SỬA: Chỉ tính điểm 1 và điểm 2 là điểm kém
                        key = f"{name_part} (Môn {mon})" if name_part else f"Môn {mon}"
                        bad_marks_list.append({'type': cat_dk, 'key': key, 'mon': mon}) # Lỗi điểm kém            
            # THUẬT TOÁN DỰ PHÒNG: Nếu không tách được theo tên, quét toàn bộ số nguyên hợp lệ trong ô
            if not parsed_any:
                # [ĐÃ VÁ LỖI]: Bắt thêm dấu + ở phía sau cho thuật toán dự phòng
                numbers = re.findall(r'(?<!\d)(\+)?\s*(10|9|8|0|[1-2])\s*(\+)?\s*(?:đ|Đ|điểm|Điểm)?(?!\d)', diem_raw)
                for has_plus_before, num_str, has_plus_after in numbers:
                    num = int(num_str)
                    
                    # [CHỐNG TRỪ ĐIỂM OAN]: Bỏ qua nếu có dấu + trước hoặc sau
                    if (has_plus_before == '+' or has_plus_after == '+') and num in [1, 2]:
                        continue
                    
                    tiet = str(df.iloc[i, 2]).strip()
                    tiet_str = f"Tiết {tiet}" if tiet != 'nan' else "Tiết học"
                    
                    if num == 10:
                        c10 += 1
                        if mon not in subject_scores: subject_scores[mon] = {'c10': 0, 'c9': 0, 'c8': 0}
                        subject_scores[mon]['c10'] += 1
                    elif num == 9:
                        c9 += 1
                        if mon not in subject_scores: subject_scores[mon] = {'c10': 0, 'c9': 0, 'c8': 0}
                        subject_scores[mon]['c9'] += 1
                    elif num == 8:
                        c8 += 1
                        if mon not in subject_scores: subject_scores[mon] = {'c10': 0, 'c9': 0, 'c8': 0}
                        subject_scores[mon]['c8'] += 1
                        
                    elif num == 0:
                        bad_marks_list.append({'type': 'Không học bài', 'key': f"Môn {mon}", 'mon': mon})
                    elif num in [1, 2]:
                        bad_marks_list.append({'type': 'Bị điểm kém', 'key': f"Môn {mon}", 'mon': mon})

        # THUẬT TOÁN XẾP LOẠI TUẦN THEO QUY CHẾ CỦA TRƯỜNG
        xep_loai = "Bình thường"
        summary_row = df[df[0].astype(str).str.contains('Tổng số giờ xếp loại', na=False, case=False)]
        if not summary_row.empty:
            sum_text = str(summary_row.iloc[0, 0])
            match_tot = re.search(r'\[\s*(\d+)\s*Tốt', sum_text, re.IGNORECASE)
            match_kha = re.search(r';\s*(\d+)\s*Khá', sum_text, re.IGNORECASE)
            match_tb  = re.search(r';\s*(\d+)\s*(?:Đạt|Trung bình|TB)', sum_text, re.IGNORECASE)
            match_yeu = re.search(r';\s*(\d+)\s*Yếu', sum_text, re.IGNORECASE)
            
            so_tot = int(match_tot.group(1)) if match_tot else 0
            so_kha = int(match_kha.group(1)) if match_kha else 0
            so_tb  = int(match_tb.group(1)) if match_tb else 0
            so_yeu = int(match_yeu.group(1)) if match_yeu else 0
            
            tong_tiet = so_tot + so_kha + so_tb + so_yeu
            
            if tong_tiet > 0:
                if so_tot == tong_tiet:
                    xep_loai = "Tuần Tốt"
                elif so_yeu == 0 and so_tb == 0 and so_kha > 0:
                    xep_loai = "Tuần Khá"

        # --- THUẬT TOÁN XÉN TRẦN LỖI ĐIỂM KÉM / KHÔNG HỌC BÀI THEO BAREM ---
        max_tot = 14
        max_mon = 4
        try:
            with session_scope() as db_session:
                active_year = db_session.query(SchoolYear).filter_by(is_active=True).first()
                if active_year:
                    settings = db_session.query(ScoreSettings).filter_by(school_year_id=active_year.id).first()
                    if settings:
                        max_tot = int(getattr(settings, 'max_diem_tot', 14))
                        max_mon = int(getattr(settings, 'max_diem_mon', 4))
        except Exception:
            pass

        # 1. Gom nhóm lỗi theo từng môn học
        bad_by_subj = {}
        for bm in bad_marks_list:
            m = bm['mon']
            if m not in bad_by_subj:
                bad_by_subj[m] = []
            bad_by_subj[m].append(bm)

        # 2. Xén bớt số lượng lỗi mỗi môn (Không vượt quá max_mon)
        surviving_bad_marks = []
        for m, marks in bad_by_subj.items():
            surviving_bad_marks.extend(marks[:max_mon])

        # 3. Xén bớt tổng số lỗi toàn tuần (Không vượt quá max_tot)
        surviving_bad_marks = surviving_bad_marks[:max_tot]

        # 4. Tái tạo lại từ điển đếm số lượng sau khi đã xén trần
        final_bad_counts = {}
        for bm in surviving_bad_marks:
            err_type = bm['type']
            key = bm['key']
            dict_key = (err_type, key)
            final_bad_counts[dict_key] = final_bad_counts.get(dict_key, 0) + 1

        # =========================================================================
        # --- [BẢN VÁ LỖI]: ĐẶT Ở ĐÂY ĐỂ ĐẢM BẢO QUÉT HẾT TUẦN MỚI BẮT ĐẦU ĐẾM ---
        # =========================================================================
        for err_type, stu_name, day in general_violations_set:
            dict_key = (err_type, stu_name)
            final_bad_counts[dict_key] = final_bad_counts.get(dict_key, 0) + 1

        # ĐÓNG GÓI LỖI VÀO Ô GHI CHÚ (SỔ ĐEN)
        note_fragments = []
        for (err_type, key), count in final_bad_counts.items():
            # ===> CHẶN TẠI ĐÂY: Nếu là lỗi "Vắng học có phép" thì BỎ QUA, không cho vào Ghi chú <===
            if err_type == 'Vắng học có phép':
                continue
                
            if key:
                note_fragments.append(f"{err_type} x{count} [{key}]")
            else:
                note_fragments.append(f"{err_type} x{count}")
        # ------------------------------------------------------------------

        # ĐÓNG GÓI MẢNG ĐIỂM MÔN HỌC CHI TIẾT
        raw_list = []
        for subj, pts in subject_scores.items():
            raw_list.append({"subj": subj, "c10": pts['c10'], "c9": pts['c9'], "c8": pts['c8']})

        return {
            "success": True,
            "branch_name": found_class_name,
            "c10": c10, "c9": c9, "c8": c8,
            "xep_loai": xep_loai,
            "so_tot": so_tot,     # <--- Bổ sung số tiết Tốt
            "so_kha": so_kha,     # <--- Bổ sung số tiết Khá
            "so_tb": so_tb,       # <--- Bổ sung số tiết Trung bình
            "so_yeu": so_yeu,     # <--- Bổ sung số tiết Yếu
            "note": " ; ".join(note_fragments),
            "raw_list": raw_list
        }
    except Exception as e:
        import traceback; traceback.print_exc()
        return {"error": f"Lỗi xử lý file Sổ Đầu Bài: {str(e)}"}, 500

@app.route('/')
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        user_data = verify_external_login(username, password)

        if user_data: 
            session['username'] = user_data.get('username', username)
            session['role'] = user_data.get('role', 'Quản trị viên')
            session['full_name'] = user_data.get('full_name', 'Người dùng')
            
            log_system_action("ĐĂNG NHẬP", f"Tài khoản {session['username']} ({session['full_name']}) đã truy cập hệ thống.")
            
            # [NÂNG CẤP LÕI]: Bẻ lái tự động theo phân quyền
            if session['role'] == "Giáo viên chủ nhiệm":
                return redirect(url_for('class_dashboard'))
            elif session['role'] == "Sao đỏ":
                return redirect(url_for('mobile_sao_do')) # Bắn thẳng vào Web App Mobile cho học sinh
            elif session['role'] == "Ban Giám hiệu":
                return redirect(url_for('bgh_dashboard')) # [MỚI]: Chuyển thẳng BGH vào Trung tâm điều hành
            else:
                return redirect(url_for('dashboard')) # Admin / Bí thư vào trang tổng quan
        else:
            flash("Tài khoản hoặc mật khẩu không chính xác!", "error")
            return redirect(url_for('login'))
            
    return render_template('login.html')

@app.route('/logout')
def logout():
    log_system_action("ĐĂNG XUẤT", f"Tài khoản {session.get('username', '')} đã rời khỏi hệ thống.")
    session.clear() 
    return redirect(url_for('login'))

# [NÂNG CẤP LÕI]: API XỬ LÝ PHÚC KHẢO TÍCH HỢP AUTO-CORRECTION, AUTO-CLEAN, PUSH NOTIFICATION & LƯU VẾT NGƯỜI DUYỆT
@app.route('/resolve_appeal', methods=['POST'])
def resolve_appeal():
    # Chống GVCN can thiệp
    if session.get('role') == 'Giáo viên chủ nhiệm':
        return redirect(url_for('login'))
        
    score_id = request.form.get('score_id', type=int)
    action = request.form.get('action') 
    response_text = request.form.get('response_text', '').strip()
    refund_points = request.form.get('refund_points', type=float, default=0.0)
    
    if not score_id or not response_text:
        flash("Vui lòng nhập nội dung phản hồi phán quyết!", "error")
        return redirect(request.referrer or url_for('dashboard'))
        
    try:
        with session_scope() as db_session:
            score = db_session.query(WeeklyScore).filter_by(id=score_id).first()
            if score:
                branch_name = score.branch.name.strip().upper() # Tên lớp dùng làm username nhận thông báo
                week_num = score.week

                # =======================================================
                # [BỔ SUNG MỚI]: Bắt thời gian VN và Người xử lý hiện tại
                # =======================================================
                from datetime import datetime, timezone, timedelta
                vn_tz = timezone(timedelta(hours=7))
                now_str = datetime.now(vn_tz).strftime("%H:%M %d/%m/%Y")
                responder = session.get('full_name', 'BCH Đoàn trường')

                # TRƯỜNG HỢP 1: ĐỒNG Ý PHÚC KHẢO & TỰ ĐỘNG TÍNH TOÁN THEO TỪNG PHẦN
                # =======================================================
                if action == 'approve':
                    auto_refund_points = 0.0
                    
                    # [BẢN VÁ LỖI 1]: Bắt mảng lỗi an toàn hơn
                    approved_errors = request.form.getlist('approved_errors[]')
                    if not approved_errors:
                        approved_errors = request.form.getlist('approved_errors')
                    
                    # Bỏ kiểm tra chuỗi cứng nhắc, chỉ cần có lỗi được tích chọn là xử lý
                    if approved_errors:
                        try:
                            
                            current_notes = [n.strip() for n in (score.note or "").split(";") if n.strip()]
                            updated_notes = [] # Danh sách mới chứa các lỗi (đã gắn tag bảo tồn)
                            
                            all_categories = db_session.query(ViolationCategory).filter_by(school_year_id=score.branch.school_year_id).all()
                            sorted_cats = sorted(all_categories, key=lambda x: len(x.name), reverse=True)
                            
                            # Quét từng lỗi đang có trong Sổ đen của lớp
                            for n in current_notes:
                                # Nếu lỗi này ĐÃ TỪNG được duyệt gỡ trước đó rồi -> Bỏ qua, giữ nguyên mộc (ĐÃ GỠ)
                                if "(ĐÃ GỠ)" in n:
                                    updated_notes.append(n)
                                    continue
                                    
                                is_approved = False
                                # [BẢN VÁ LỖI 2]: Là phẳng mọi khoảng trắng và đưa về chữ thường để so khớp chính xác tuyệt đối
                                n_clean_lower = n.lower().replace(" ", "")
                                
                                for app_err in approved_errors:
                                    # Loại bỏ phần đuôi "(Phạt Xđ)"
                                    app_err_clean = re.sub(r'\(Phạt .*?đ\)', '', app_err).strip()
                                    app_err_clean_lower = app_err_clean.lower().replace(" ", "")
                                    
                                    # So khớp an toàn tuyệt đối
                                    if app_err_clean_lower in n_clean_lower or n_clean_lower in app_err_clean_lower:
                                        is_approved = True
                                        break
                                        
                                if is_approved:
                                    # [TÍNH NĂNG MỚI]: KHÔNG XÓA LỖI ĐI MÀ ĐÓNG DẤU "(ĐÃ GỠ)" ĐỂ LƯU LỊCH SỬ
                                    updated_notes.append(f"{n} (ĐÃ GỠ)")
                                    
                                    # LỖI ĐƯỢC DUYỆT GỠ: Trích xuất cộng điểm hoàn trả
                                    day_pfx_match = re.search(r'\[(T[2-7](?:\s*Chiều|\s*Chieu)?|CN)\]', n, re.IGNORECASE)
                                    day_pfx = day_pfx_match.group(0) if day_pfx_match else ""
                                    text_to_parse = n.replace(day_pfx, "").strip() if day_pfx else n
                                    
                                    for cat in sorted_cats:
                                        if cat.name.lower() in text_to_parse.lower() and getattr(cat, 'point_type', 'Điểm trừ') != 'Điểm cộng':
                                            qty_match = re.search(r'(?:x|:|-)\s*(\d+)', text_to_parse.lower())
                                            qty = int(qty_match.group(1)) if qty_match else 1
                                            auto_refund_points += float(cat.penalty_points * qty)
                                            break
                                else:
                                    # LỖI BỊ TỪ CHỐI HOẶC KHÔNG KHIẾU NẠI -> Giữ lại trong Sổ đen
                                    updated_notes.append(n)
                                    
                            # Ghi đè lại ghi chú (Đã bảo toàn 100% dữ liệu gốc, chỉ thêm chữ ĐÃ GỠ)
                            score.note = " ; ".join(updated_notes)
                            
                            # ĐỒNG BỘ LÀM SẠCH "SỔ ĐEN TOÀN TRƯỜNG" DỰA TRÊN PHẦN CÒN LẠI
                            db_session.query(WeeklyViolation).filter_by(weekly_score_id=score.id).delete()
                            
                            for part in updated_notes:
                                # [QUAN TRỌNG]: Lỗi nào có mộc (ĐÃ GỠ) thì loại thẳng tay khỏi Sổ đen!
                                if "(ĐÃ GỠ)" in part: 
                                    continue 
                                    
                                match_day = re.search(r'\[(T[2-7](?:\s*Chiều|\s*Chieu)?|CN)\]', part, re.IGNORECASE)
                                day_pfx = match_day.group(0) if match_day else ""
                                text_to_parse = part.replace(day_pfx, "").strip() if day_pfx else part
                                
                                match_stu = re.search(r'\[(.*?)\]', text_to_parse)
                                stu_display = match_stu.group(1).strip() if match_stu else None
                                
                                for cat in sorted_cats:
                                    if cat.name.lower() in text_to_parse.lower():
                                        match_qty = re.search(r'(?:x|:|-)\s*(\d+)', text_to_parse.lower())
                                        qty = int(match_qty.group(1)) if match_qty else 1
                                        db_session.add(WeeklyViolation(weekly_score_id=score.id, violation_id=cat.id, quantity=qty, student_name=stu_display))
                                        break
                        except Exception as e:
                            print(f"Lỗi tự động xóa Sổ đen: {e}")
                    
                    # Ưu tiên lấy điểm tự động tính toán bởi Python Backend.
                    form_refund = request.form.get('refund_points', type=float, default=0.0)
                    final_refund = auto_refund_points if auto_refund_points > 0 else form_refund
                    
                    # Thực hiện hoàn điểm
                    score.total_score = float(score.total_score or 0) + final_refund
                    score.score_tru = max(0.0, float(score.score_tru or 0) - final_refund)
                    
                    # ====================================================================
                    # [NÂNG CẤP LÕI]: Đóng dấu "Người xử lý" và "Thời gian" vào chuỗi
                    # ====================================================================
                    score.appeal_response = f"[ĐÃ DUYỆT] Đã gỡ lỗi được chọn và hoàn {final_refund}đ. Phản hồi: {response_text}\n(Xử lý bởi: {responder} lúc {now_str})"
                    
                    log_system_action("XỬ LÝ PHÚC KHẢO", f"Đã DUYỆT 1 PHẦN khiếu nại lớp {score.branch.name} Tuần {score.week}. Tự động hoàn {final_refund}đ.")
                    flash(f"✅ Đã duyệt khiếu nại, hệ thống hoàn {final_refund}đ và xử lý Sổ đen chuẩn xác!", "success")
                    
                    # [NÂNG CẤP]: BẮN THÔNG BÁO ĐẨY CHO GVCN KHI ĐƯỢC DUYỆT PHÚC KHẢO
                    try:
                        push_title = f"🎉 Phúc khảo {week_num} đã được DUYỆT!"
                        push_body = f"Được hoàn {final_refund}đ. Phản hồi: {response_text[:40]}..."
                        send_web_push(branch_name, push_title, push_body)
                    except Exception as err:
                        print(f"Lỗi gửi Push thông báo duyệt phúc khảo: {err}")
                
                # =======================================================
                # TRƯỜNG HỢP 2: TỪ CHỐI PHÚC KHẢO
                # =======================================================
                elif action == 'reject':
                    # ====================================================================
                    # [NÂNG CẤP LÕI]: Đóng dấu "Người xử lý" và "Thời gian" vào chuỗi
                    # ====================================================================
                    score.appeal_response = f"[TỪ CHỐI] Phản hồi: {response_text}\n(Xử lý bởi: {responder} lúc {now_str})"
                    
                    log_system_action("XỬ LÝ PHÚC KHẢO", f"TỪ CHỐI khiếu nại lớp {score.branch.name} Tuần {score.week}: {response_text}")
                    flash(f"Đã đóng Ticket và từ chối khiếu nại của lớp {score.branch.name}.", "warning")
                    
                    # [NÂNG CẤP]: BẮN THÔNG BÁO ĐẨY CHO GVCN KHI BỊ TỪ CHỐI PHÚC KHẢO
                    try:
                        push_title = f"📢 Phản hồi phúc khảo {week_num}"
                        push_body = f"Yêu cầu chưa được chấp thuận. Phản hồi: {response_text[:40]}..."
                        send_web_push(branch_name, push_title, push_body)
                    except Exception as err:
                        print(f"Lỗi gửi Push thông báo từ chối phúc khảo: {err}")
                    
    except Exception as e:
        import traceback; traceback.print_exc()
        flash(f"Lỗi xử lý phúc khảo: {e}", "error")
        
    return redirect(request.referrer or url_for('dashboard'))

@app.route('/dashboard')
def dashboard():
    try:
        with session_scope() as db_session:
            from sqlalchemy import func
            import re
            active_year = db_session.query(SchoolYear).filter(SchoolYear.is_active == True).first()
            
            total_branches = 0
            current_week = "Chưa có"
            avg_score = 0.0
            chart_labels = []
            chart_data = []
            pending_appeals_data = []
            
            # [TÍNH NĂNG MỚI]: BẢNG BÁO ĐỘNG HỌC SINH CÁ BIỆT TOÀN TRƯỜNG
            global_warnings = []

            if active_year:
                total_branches = db_session.query(Branch).filter(Branch.school_year_id == active_year.id).count()
                
                # =====================================================================
                # [BẢN VÁ LỖI]: Tìm chính xác số tuần lớn nhất (Toán học) thay vì thứ tự nhập
                # =====================================================================
                all_weeks = db_session.query(WeeklyScore.week).join(Branch).filter(Branch.school_year_id == active_year.id).distinct().all()
                
                if all_weeks:
                    current_week = max([w[0] for w in all_weeks], key=lambda x: int(re.search(r'\d+', str(x)).group()) if re.search(r'\d+', str(x)) else 0)
                    
                    avg_val = db_session.query(func.avg(WeeklyScore.total_score)).join(Branch).filter(
                        Branch.school_year_id == active_year.id,
                        WeeklyScore.week == current_week
                    ).scalar()
                    
                    try: avg_score = round(float(avg_val), 1) if avg_val is not None else 0.0
                    except: avg_score = 0.0

                    top_10 = db_session.query(Branch.name, WeeklyScore.total_score).join(WeeklyScore).filter(
                        Branch.school_year_id == active_year.id,
                        WeeklyScore.week == current_week
                    ).order_by(WeeklyScore.total_score.desc()).limit(10).all()

                    chart_labels = [item[0] for item in top_10]
                    chart_data = [round(float(item[1]), 1) if item[1] is not None else 0.0 for item in top_10]

                if session.get('role') != 'Giáo viên chủ nhiệm':
                    appeals = db_session.query(WeeklyScore).join(Branch).filter(
                        Branch.school_year_id == active_year.id,
                        WeeklyScore.is_appealed == True,
                        (WeeklyScore.appeal_response == None) | (WeeklyScore.appeal_response == "")
                    ).all()
                    for p in appeals:
                        pending_appeals_data.append({
                            'score_id': p.id, 
                            'branch_name': p.branch.name,
                            'week': p.week, 
                            'reason': p.appeal_reason,
                            'note': p.note  # <--- DÒNG KÉO DỮ LIỆU BGH VẪN ĐƯỢC GIỮ NGUYÊN
                        })

            return render_template(
                'dashboard.html',
                user_fullname=session.get('full_name', "Người dùng"),
                user_role=session.get('role', "Quản trị viên"),
                total_branches=total_branches,
                current_week=current_week,
                avg_score=avg_score,
                chart_labels=chart_labels,
                chart_data=chart_data,
                pending_appeals=pending_appeals_data,
                global_warnings=global_warnings # Truyền biến báo động ra giao diện
            )
    except Exception as e:
        import traceback; traceback.print_exc()
        flash(f"Lỗi tải bảng điều khiển: {e}", "error")
        return redirect(url_for('login'))
    
# ==========================================
# MODULE: ĐỔI MẬT KHẨU TRỰC TUYẾN (ÉP ĐỔI)
# ==========================================
@app.route('/change-password', methods=['GET', 'POST'])
def change_password():
    if 'username' not in session:
        flash("Vui lòng đăng nhập trước khi đổi mật khẩu!", "warning")
        return redirect(url_for('login'))

    if request.method == 'POST':
        login_id = request.form.get('login_id', '').strip()
        new_pwd = request.form.get('new_password', '').strip()
        confirm_pwd = request.form.get('confirm_password', '').strip()

        if not login_id or not new_pwd or not confirm_pwd:
            flash("Vui lòng nhập đầy đủ thông tin Tên đăng nhập và Mật khẩu mới!", "error")
            return redirect(url_for('change_password'))

        if len(new_pwd) < 3:
            flash("Mật khẩu mới phải có ít nhất 3 ký tự!", "error")
            return redirect(url_for('change_password'))

        if new_pwd != confirm_pwd:
            flash("Mật khẩu xác nhận không khớp!", "error")
            return redirect(url_for('change_password'))

        try:
            accounts = load_external_accounts()
            target_acc = None
            for acc in accounts:
                if acc.get("username") == login_id:
                    target_acc = acc
                    break

            if not target_acc:
                flash(f"Không tìm thấy tài khoản nào có Tên đăng nhập là '{login_id}'!", "error")
                return redirect(url_for('change_password'))

            target_acc["password"] = new_pwd
            save_external_accounts(accounts)

            with session_scope() as session_db:
                db_user = session_db.query(User).filter(User.username == login_id).first()
                if db_user:
                    db_user.password_hash = new_pwd

            log_system_action("BẢO MẬT", f"Đã đổi mật khẩu cho tài khoản: {login_id}")

            flash(f"✅ Đã thay đổi mật khẩu thành công cho tài khoản: {login_id}! Vui lòng đăng nhập lại.", "success")
            session.clear() 
            return redirect(url_for('login'))

        except Exception as e:
            flash(f"Lỗi hệ thống khi đổi mật khẩu: {str(e)}", "error")
            return redirect(url_for('change_password'))

    return render_template('change_password.html', default_username=session.get('username', ''))

# ==========================================
# MODULE: XEM NHẬT KÝ HỆ THỐNG (CHỈ DÀNH CHO ADMIN)
# ==========================================
@app.route('/action_logs')
def action_logs():
    user_role = session.get('role', '')
    
    if user_role not in ['Quản trị viên', 'Admin', 'Bí thư', 'Bí thư Đoàn trường']:
        flash("Bạn không có quyền xem Nhật ký hệ thống!", "error")
        return redirect(url_for('dashboard'))
        
    try:
        import datetime as dt
        # Lấy từ khóa tài khoản cần lọc từ URL (nếu có)
        search_user = request.args.get('search_user', '').strip()
        
        with session_scope() as db_session:
            query = db_session.query(ActionLog)
            
            # Nếu có nhập tên tài khoản, tiến hành lọc gần đúng (case-insensitive)
            if search_user:
                query = query.filter(ActionLog.username.ilike(f"%{search_user}%"))
                
            logs = query.order_by(ActionLog.timestamp.desc()).limit(500).all()
            
            logs_data = []
            for log in logs:
                local_time = log.timestamp + dt.timedelta(hours=7) if log.timestamp else dt.datetime.now()
                logs_data.append({
                    'time': local_time.strftime("%d/%m/%Y - %H:%M:%S"),
                    'username': log.username,
                    'full_name': log.full_name,
                    'action_type': log.action_type,
                    'details': log.details
                })
                
        return render_template('action_logs.html', logs=logs_data, search_user=search_user)
    except Exception as e:
        flash(f"Lỗi tải nhật ký: {e}", "error")
        return redirect(url_for('dashboard'))

# ==========================================
# CẤU HÌNH HỆ THỐNG: QUẢN LÝ NĂM HỌC
# ==========================================
@app.route('/school-years', methods=['GET'])
def school_years():
    try:
        with session_scope() as db_session:
            years = db_session.query(SchoolYear).order_by(SchoolYear.id.desc()).all()
            return render_template('school_years.html', years=years)
    except Exception as e:
        flash(f"Lỗi tải danh sách năm học: {e}", "error")
        return redirect(url_for('dashboard'))

@app.route('/add_school_year', methods=['POST'])
def add_school_year():
    name = request.form.get('name', '').strip()
    if not name:
        flash("Vui lòng điền tên năm học!", "error")
        return redirect(url_for('school_years'))

    try:
        with session_scope() as db_session:
            is_first = db_session.query(SchoolYear).count() == 0
            new_year = SchoolYear(name=name, is_active=is_first)
            db_session.add(new_year)
            log_system_action("CẤU HÌNH", f"Thêm năm học mới: {name}")
        flash(f"Đã thêm năm học mới: {name}", "success")
    except Exception as e:
        flash(f"Lỗi khi thêm năm học: {e}", "error")
    return redirect(url_for('school_years'))

@app.route('/set_active_year/<int:id>', methods=['POST'])
def set_active_year(id):
    try:
        with session_scope() as db_session:
            db_session.query(SchoolYear).update({SchoolYear.is_active: False})
            selected_year = db_session.query(SchoolYear).filter_by(id=id).first()
            if selected_year:
                selected_year.is_active = True
                log_system_action("CẤU HÌNH", f"Kích hoạt năm học: {selected_year.name}")
                flash(f"Đã chuyển sang năm học: {selected_year.name}", "success")
    except Exception as e:
        flash(f"Lỗi kích hoạt năm học: {e}", "error")
    return redirect(url_for('school_years'))

@app.route('/delete_school_year/<int:id>', methods=['POST'])
def delete_school_year(id):
    try:
        with session_scope() as db_session:
            year = db_session.query(SchoolYear).filter_by(id=id).first()
            if year:
                if year.is_active:
                    flash("Không thể xóa năm học đang được kích hoạt!", "error")
                else:
                    name = year.name
                    db_session.delete(year)
                    log_system_action("CẤU HÌNH", f"Xóa năm học: {name}")
                    flash("Đã xóa năm học thành công!", "success")
    except Exception as e:
        flash(f"Lỗi khi xóa năm học: {e}", "error")
    return redirect(url_for('school_years'))


# ==========================================
# CẤU HÌNH HỆ THỐNG: QUẢN LÝ NGƯỜI DÙNG
# ==========================================
@app.route('/users', methods=['GET'])
def users():
    try:
        with session_scope() as db_session:
            user_list = db_session.query(User).order_by(User.id).all()
            return render_template('users.html', users=user_list)
    except Exception as e:
        flash(f"Lỗi tải danh sách người dùng: {e}", "error")
        return redirect(url_for('dashboard'))

@app.route('/add_user', methods=['POST'])
def add_user():
    username = request.form.get('username', '').strip()
    password = request.form.get('password', '').strip()
    full_name = request.form.get('full_name', '').strip()
    role_text = request.form.get('role', '').strip()

    if not username or not full_name or not password:
        flash("Vui lòng điền đầy đủ Tên đăng nhập, Họ tên và Mật khẩu!", "error")
        return redirect(url_for('users'))

    try:
        with session_scope() as db_session:
            # [ĐÃ SỬA]: Bổ sung nhận diện đầy đủ các quyền, đặc biệt là Ban Giám hiệu
            role_enum = UserRole.BCH
            if "Quản trị" in role_text or "Admin" in role_text: 
                role_enum = UserRole.ADMIN
            elif "Bí thư" in role_text: 
                role_enum = UserRole.BI_THU
            elif "Giáo viên chủ nhiệm" in role_text: 
                role_enum = UserRole.GVCN
            elif "Sao đỏ" in role_text: 
                role_enum = UserRole.SAO_DO
            elif "Ban Giám hiệu" in role_text or "BGH" in role_text: 
                role_enum = getattr(UserRole, 'BGH', getattr(UserRole, 'BAN_GIAM_HIEU', UserRole.BCH))

            exist = db_session.query(User).filter_by(username=username).first()
            if exist:
                flash(f"Tên đăng nhập '{username}' đã tồn tại trên hệ thống!", "error")
            else:
                new_user = User(username=username, password_hash=password, full_name=full_name, role=role_enum, is_active=True)
                db_session.add(new_user)
                
                try: sync_account_to_json(username, full_name, password, role_text, True)
                except Exception as e: print(f"Lỗi đồng bộ JSON: {e}")
                
                log_system_action("TẠO TÀI KHOẢN", f"Đã cấp tài khoản mới: {username} (Họ tên: {full_name}, Quyền: {role_text})")
                flash(f"Đã cấp tài khoản thành công cho: {full_name}", "success")
    except Exception as e:
        flash(f"Lỗi khi thêm tài khoản: {e}", "error")
    return redirect(url_for('users'))

@app.route('/edit_user/<int:id>', methods=['POST'])
def edit_user(id):
    new_username = request.form.get('edit_username', '').strip()
    new_fullname = request.form.get('edit_fullname', '').strip()
    new_password = request.form.get('edit_password', '').strip()
    new_role_text = request.form.get('edit_role', '').strip()

    if not new_username or not new_fullname:
        flash("Tên đăng nhập và Họ tên không được để trống!", "error")
        return redirect(url_for('users'))

    try:
        with session_scope() as db_session:
            user = db_session.query(User).filter_by(id=id).first()
            if not user:
                flash("Không tìm thấy tài khoản này!", "error")
                return redirect(url_for('users'))

            old_username = user.username
            changes_made = False
            change_details = []

            if new_username != old_username:
                exist = db_session.query(User).filter_by(username=new_username).first()
                if exist:
                    flash(f"Tên đăng nhập '{new_username}' đã có người sử dụng!", "error")
                    return redirect(url_for('users'))
                user.username = new_username
                changes_made = True
                change_details.append(f"Tên ĐN: '{old_username}' -> '{new_username}'")

            if new_fullname != user.full_name:
                change_details.append(f"Họ tên: '{user.full_name}' -> '{new_fullname}'")
                user.full_name = new_fullname
                changes_made = True

            # [ĐÃ SỬA]: Bổ sung nhận diện phân quyền khi chỉnh sửa
            new_role_enum = UserRole.BCH
            if "Quản trị" in new_role_text or "Admin" in new_role_text: 
                new_role_enum = UserRole.ADMIN
            elif "Bí thư" in new_role_text: 
                new_role_enum = UserRole.BI_THU
            elif "Giáo viên chủ nhiệm" in new_role_text: 
                new_role_enum = UserRole.GVCN
            elif "Sao đỏ" in new_role_text: 
                new_role_enum = UserRole.SAO_DO
            elif "Ban Giám hiệu" in new_role_text or "BGH" in new_role_text: 
                new_role_enum = getattr(UserRole, 'BGH', getattr(UserRole, 'BAN_GIAM_HIEU', UserRole.BCH))
                
            if user.role != new_role_enum:
                change_details.append(f"Quyền: '{user.role.value}' -> '{new_role_text}'")
                user.role = new_role_enum
                changes_made = True

            if new_password:
                if len(new_password) < 3:
                    flash("Mật khẩu mới phải có ít nhất 3 ký tự!", "error")
                    return redirect(url_for('users'))
                user.password_hash = new_password
                changes_made = True
                change_details.append("Đã đổi mật khẩu")

            if changes_made:
                try:
                    if new_username != old_username: remove_account_from_json(old_username)
                    sync_account_to_json(user.username, user.full_name, user.password_hash, new_role_text, user.is_active)
                except: pass

                log_system_action("CHỈNH SỬA TÀI KHOẢN", f"Cập nhật '{old_username}': " + "; ".join(change_details))
                flash(f"Đã cập nhật thông tin tài khoản '{new_username}'!", "success")
            else:
                flash("Chưa có thông tin nào được thay đổi.", "info")

    except Exception as e:
        flash(f"Lỗi cập nhật tài khoản: {e}", "error")
    return redirect(url_for('users'))
# ==========================================
# API: THAO TÁC HÀNG LOẠT VỚI TÀI KHOẢN (XÓA / KHÓA)
# ==========================================
@app.route('/bulk_user_action', methods=['POST'])
def bulk_user_action():
    # Chỉ Admin/Bí thư mới được làm điều này
    if session.get('role') not in ['Quản trị viên', 'Admin', 'Bí thư Đoàn trường', 'Bí thư']:
        flash("Bạn không có quyền thực hiện chức năng này!", "error")
        return redirect(url_for('users'))

    action = request.form.get('action') # Nhận lệnh: 'delete' (xóa) hoặc 'toggle_lock' (khóa/mở)
    user_ids = request.form.getlist('user_ids') # Lấy danh sách các ID được tích chọn

    if not user_ids:
        flash("Vui lòng tích chọn ít nhất 1 tài khoản để thao tác!", "warning")
        return redirect(url_for('users'))

    try:
        with session_scope() as db_session:
            count = 0
            for uid in user_ids:
                user = db_session.query(User).filter_by(id=int(uid)).first()
                if not user:
                    continue

                # CƠ CHẾ BẢO VỆ LÕI: Tuyệt đối không cho chạm vào tài khoản admin tối cao
                if user.username.lower() == 'admin':
                    continue

                if action == 'delete':
                    name = user.username
                    db_session.delete(user)
                    try: remove_account_from_json(name)
                    except: pass
                    count += 1
                    
                elif action == 'toggle_lock':
                    user.is_active = not user.is_active
                    role_text = user.role.value if hasattr(user.role, 'value') else str(user.role)
                    try: sync_account_to_json(user.username, user.full_name, user.password_hash, role_text, user.is_active)
                    except: pass
                    count += 1

            # Phản hồi theo từng hành động
            if action == 'delete':
                log_system_action("XÓA TÀI KHOẢN", f"Đã xóa hàng loạt {count} tài khoản khỏi hệ thống.")
                flash(f"✅ Đã xóa vĩnh viễn {count} tài khoản thành công!", "success")
            elif action == 'toggle_lock':
                log_system_action("THAY ĐỔI TRẠNG THÁI", f"Đã thay đổi trạng thái hàng loạt {count} tài khoản.")
                flash(f"✅ Đã Khóa / Mở khóa {count} tài khoản thành công!", "success")

    except Exception as e:
        import traceback
        traceback.print_exc()
        flash(f"Lỗi thao tác hàng loạt: {str(e)}", "error")

    return redirect(url_for('users'))

@app.route('/toggle_user/<int:id>', methods=['POST'])
def toggle_user(id):
    try:
        with session_scope() as db_session:
            user = db_session.query(User).filter_by(id=id).first()
            if user:
                if user.username.lower() == 'admin':
                    flash("Không thể khóa tài khoản quản trị viên tối cao!", "error")
                else:
                    user.is_active = not user.is_active
                    status = "Mở khóa" if user.is_active else "Khóa"
                    
                    try: sync_account_to_json(user.username, user.full_name, user.password_hash, user.role.value, user.is_active)
                    except: pass
                    
                    log_system_action("THAY ĐỔI TRẠNG THÁI", f"Đã {status} tài khoản: {user.username}")
                    flash(f"Đã {status} tài khoản: {user.username}", "success")
    except Exception as e:
        flash(f"Lỗi thay đổi trạng thái user: {e}", "error")
    return redirect(url_for('users'))

@app.route('/delete_user/<int:id>', methods=['POST'])
def delete_user(id):
    try:
        with session_scope() as db_session:
            user = db_session.query(User).filter_by(id=id).first()
            if user:
                if user.username.lower() == 'admin':
                    flash("Hệ thống từ chối xóa tài khoản quản trị tối cao!", "error")
                else:
                    name = user.username
                    db_session.delete(user)
                    
                    try: remove_account_from_json(name)
                    except: pass
                    
                    log_system_action("XÓA TÀI KHOẢN", f"Đã xóa vĩnh viễn tài khoản: {name}")
                    flash(f"Đã xóa vĩnh viễn tài khoản: {name}", "success")
    except Exception as e:
        flash(f"Lỗi xóa tài khoản: {e}", "error")
    return redirect(url_for('users'))

# ==========================================
# MODULE: QUẢN LÝ CHI ĐOÀN
# ==========================================
@app.route('/branches', methods=['GET', 'POST'])
def branches():
    try:
        with session_scope() as db_session:
            active_year = db_session.query(SchoolYear).filter(SchoolYear.is_active == True).first()
            
            if request.method == 'POST':
                if not active_year:
                    flash("Chưa có năm học nào được kích hoạt!", "error")
                    return redirect(url_for('branches'))
                    
                name = request.form.get('name', '').strip().upper()
                group = request.form.get('group', 'Nhóm 1')
                si_so = request.form.get('si_so', 0)
                gvcn = request.form.get('gvcn', '').strip()
                phone_gvcn = request.form.get('phone_gvcn', '').strip()
                class_monitor = request.form.get('class_monitor', '').strip()
                phone_monitor = request.form.get('phone_monitor', '').strip()
                
                try: si_so = int(si_so) if si_so else 0
                except ValueError: si_so = 0

                if name:
                    exist = db_session.query(Branch).filter(Branch.name == name, Branch.school_year_id == active_year.id).first()
                    if exist:
                        flash(f"Chi đoàn {name} đã tồn tại trong năm học này!", "error")
                    else:
                        new_branch = Branch(
                            name=name, group=group, si_so=si_so, gvcn=gvcn, 
                            phone_gvcn=phone_gvcn, class_monitor=class_monitor, phone_monitor=phone_monitor,
                            school_year_id=active_year.id
                        )
                        db_session.add(new_branch)
                        log_system_action("QUẢN LÝ CHI ĐOÀN", f"Đã thêm Chi đoàn {name}")
                        flash(f"Đã thêm Chi đoàn {name} thành công!", "success")
                else:
                    flash("Vui lòng nhập tên Chi đoàn!", "error")
                    
                return redirect(url_for('branches'))

            branch_list = []
            if active_year:
                branch_list = db_session.query(Branch).filter(Branch.school_year_id == active_year.id).all()
                
            return render_template('branches.html', branches=branch_list, active_year=active_year)
    except Exception as e:
        flash(f"Lỗi phân hệ Chi đoàn: {e}", "error")
        return redirect(url_for('dashboard'))

@app.route('/delete_branch/<int:id>', methods=['POST'])
def delete_branch(id):
    try:
        with session_scope() as db_session:
            branch = db_session.query(Branch).filter(Branch.id == id).first()
            if branch:
                name = branch.name
                db_session.delete(branch)
                log_system_action("QUẢN LÝ CHI ĐOÀN", f"Đã xóa Chi đoàn {name}")
                flash(f"Đã xóa Chi đoàn {name} thành công!", "success")
    except Exception as e:
        flash(f"Lỗi xóa Chi đoàn: {e}", "error")
    return redirect(url_for('branches'))

@app.route('/edit_branch/<int:id>', methods=['POST'])
def edit_branch(id):
    try:
        with session_scope() as db_session:
            branch = db_session.query(Branch).filter(Branch.id == id).first()
            if branch:
                branch.name = request.form.get('edit_name', branch.name).strip().upper()
                branch.group = request.form.get('edit_group', branch.group)
                branch.gvcn = request.form.get('edit_gvcn', branch.gvcn).strip()
                branch.phone_gvcn = request.form.get('edit_phone_gvcn', branch.phone_gvcn).strip()
                branch.class_monitor = request.form.get('edit_class_monitor', branch.class_monitor).strip()
                branch.phone_monitor = request.form.get('edit_phone_monitor', branch.phone_monitor).strip()
                
                si_so = request.form.get('edit_si_so', branch.si_so)
                try: branch.si_so = int(si_so) if si_so else 0
                except ValueError: branch.si_so = 0
                
                log_system_action("QUẢN LÝ CHI ĐOÀN", f"Đã cập nhật thông tin Chi đoàn {branch.name}")
                flash(f"Đã cập nhật thông tin Chi đoàn {branch.name}!", "success")
    except Exception as e:
        flash(f"Lỗi cập nhật Chi đoàn: {e}", "error")
    return redirect(url_for('branches'))

@app.route('/import_branches', methods=['POST'])
def import_branches():
    if 'excel_file' not in request.files:
        flash("Không tìm thấy file tải lên!", "error")
        return redirect(url_for('branches'))
    
    file = request.files['excel_file']
    if file.filename == '':
        flash("Chưa chọn file nào!", "error")
        return redirect(url_for('branches'))
        
    if file and (file.filename.endswith('.xlsx') or file.filename.endswith('.xls')):
        try:
            df = pd.read_excel(file)
            with session_scope() as db_session:
                active_year = db_session.query(SchoolYear).filter(SchoolYear.is_active == True).first()
                if not active_year:
                    flash("Chưa có năm học nào được kích hoạt!", "error")
                    return redirect(url_for('branches'))
                    
                count_success = 0
                for index, row in df.iterrows():
                    name = str(row.get('Tên Lớp', '')).strip().upper()
                    if not name or name == 'NAN': continue
                    
                    group = str(row.get('Nhóm', 'Nhóm 1')).strip()
                    gvcn = str(row.get('GVCN', '')).strip()
                    phone_gvcn = str(row.get('SĐT GVCN', '')).strip()
                    class_monitor = str(row.get('Lớp trưởng', '')).strip()
                    phone_monitor = str(row.get('SĐT Lớp trưởng', '')).strip()
                    
                    if gvcn.lower() == 'nan': gvcn = ''
                    if phone_gvcn.lower() == 'nan': phone_gvcn = ''
                    if class_monitor.lower() == 'nan': class_monitor = ''
                    if phone_monitor.lower() == 'nan': phone_monitor = ''
                    
                    try: si_so = int(row.get('Sĩ số', 0))
                    except: si_so = 0
                        
                    exist = db_session.query(Branch).filter(Branch.name == name, Branch.school_year_id == active_year.id).first()
                    if not exist:
                        new_branch = Branch(
                            name=name, group=group, si_so=si_so, gvcn=gvcn, 
                            phone_gvcn=phone_gvcn, class_monitor=class_monitor, phone_monitor=phone_monitor,
                            school_year_id=active_year.id
                        )
                        db_session.add(new_branch)
                        count_success += 1
                        
                log_system_action("QUẢN LÝ CHI ĐOÀN", f"Đã nhập thành công {count_success} Chi đoàn từ file Excel")
                flash(f"Đã nhập thành công {count_success} Chi đoàn từ file Excel!", "success")
        except Exception as e:
            flash(f"Lỗi đọc file Excel: {str(e)}", "error")
    else:
        flash("Vui lòng chọn file Excel hợp lệ (.xlsx, .xls)", "error")
        
    return redirect(url_for('branches'))

# ==========================================
# MODULE: QUẢN LÝ ĐỘI SAO ĐỎ
# ==========================================
@app.route('/red-stars', methods=['GET', 'POST'])
def red_stars():
    try:
        with session_scope() as db_session:
            active_year = db_session.query(SchoolYear).filter(SchoolYear.is_active == True).first()
            
            if request.method == 'POST':
                if not active_year:
                    flash("Chưa có năm học nào được kích hoạt!", "error")
                    return redirect(url_for('red_stars'))
                    
                full_name = request.form.get('full_name', '').strip()
                gender = request.form.get('gender', 'Nam')
                phone = request.form.get('phone', '').strip()
                branch_id = request.form.get('branch_id')
                notes = request.form.get('notes', '').strip()
                
                if full_name and branch_id:
                    new_star = RedStar(
                        full_name=full_name, gender=gender, phone=phone,
                        branch_id=branch_id, notes=notes, is_active=True
                    )
                    db_session.add(new_star)
                    log_system_action("ĐỘI SAO ĐỎ", f"Đã thêm Sao đỏ {full_name}")
                    flash(f"Đã thêm Sao đỏ {full_name} thành công!", "success")
                else:
                    flash("Vui lòng nhập tên và chọn Chi đoàn!", "error")
                return redirect(url_for('red_stars'))

            stars_list = []
            branches_list = []
            search_name = request.args.get('search_name', '').strip()
            filter_branch = request.args.get('filter_branch', '')
            
            # [TÍNH NĂNG MỚI]: BẢN ĐỒ LỊCH TRỰC TUẦN HIỆN TẠI
            current_assignments_map = {}

            if active_year:
                branches_list = db_session.query(Branch).filter(Branch.school_year_id == active_year.id).all()
                branch_ids = [b.id for b in branches_list]
                
                if branch_ids:
                    query = db_session.query(RedStar).filter(RedStar.branch_id.in_(branch_ids))
                    if search_name:
                        query = query.filter(RedStar.full_name.ilike(f"%{search_name}%"))
                    if filter_branch:
                        query = query.filter(RedStar.branch_id == int(filter_branch))
                    stars_list = query.all()
                    
                    # 1. Tìm tuần trực mới nhất
                    latest_assign = db_session.query(Assignment).order_by(Assignment.week_number.desc()).first()
                    current_week_num = latest_assign.week_number if latest_assign else 0
                    
                    # 2. Đọc Sơ đồ Cụm (class_zones.json) để dịch tên Khu vực ra Các lớp cụ thể
                    import os, json
                    zones_map = {}
                    config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config", "class_zones.json")
                    if os.path.exists(config_path):
                        with open(config_path, "r", encoding="utf-8") as f:
                            try: zones_map = json.load(f)
                            except: pass
                            
                    # 3. Quét lịch trực tuần hiện tại và đóng gói dữ liệu cho từng Sao đỏ
                    if current_week_num > 0:
                        current_assigns = db_session.query(Assignment).filter(Assignment.week_number == current_week_num).all()
                        for a in current_assigns:
                            if a.duty_area:
                                area_name = a.duty_area.name
                                classes = zones_map.get(area_name, [])
                                target_str = ", ".join(classes) if classes else "Khu vực chung"
                                
                                # Tạo chuỗi hiển thị HTML (Có icon và màu sắc bắt mắt)
                                info_str = f"<div class='mt-1 p-2 rounded bg-primary bg-opacity-10 border border-primary border-opacity-25' style='font-size: 12px;'>" \
                                           f"<span class='text-primary fw-bold'><i class='fa-solid fa-location-dot me-1'></i>Tuần {current_week_num}: {area_name}</span><br>" \
                                           f"<span class='text-secondary fw-bold'><i class='fa-solid fa-school me-1'></i>Chấm: {target_str}</span>" \
                                           f"</div>"
                                           
                                # Nếu một Sao đỏ trực 2 ca/tuần, sẽ cộng dồn chuỗi lại
                                if a.red_star_id not in current_assignments_map:
                                    current_assignments_map[a.red_star_id] = info_str
                                else:
                                    current_assignments_map[a.red_star_id] += info_str
                    
            return render_template(
                'red_stars.html', stars=stars_list, branches=branches_list, 
                active_year=active_year, search_name=search_name, filter_branch=filter_branch,
                current_assignments_map=current_assignments_map # Truyền bản đồ lịch trực ra Giao diện
            )
    except Exception as e:
        import traceback; traceback.print_exc()
        flash(f"Lỗi phân hệ Sao đỏ: {e}", "error")
        return redirect(url_for('dashboard'))

@app.route('/toggle_star_status/<int:id>', methods=['POST'])
def toggle_star_status(id):
    try:
        with session_scope() as db_session:
            star = db_session.query(RedStar).filter(RedStar.id == id).first()
            if star:
                star.is_active = not star.is_active
                status_label = "Đang hoạt động" if star.is_active else "Tạm nghỉ"
                log_system_action("ĐỘI SAO ĐỎ", f"Đã chuyển trạng thái của {star.full_name} sang: {status_label}")
                flash(f"Đã chuyển trạng thái của {star.full_name} sang: {status_label}", "success")
    except Exception as e:
        flash(f"Lỗi đổi trạng thái sao đỏ: {e}", "error")
    return redirect(url_for('red_stars'))

@app.route('/edit_red_star/<int:id>', methods=['POST'])
def edit_red_star(id):
    try:
        with session_scope() as db_session:
            star = db_session.query(RedStar).filter(RedStar.id == id).first()
            if star:
                star.full_name = request.form.get('edit_full_name', star.full_name).strip()
                star.gender = request.form.get('edit_gender', star.gender)
                star.phone = request.form.get('edit_phone', star.phone).strip()
                star.notes = request.form.get('edit_notes', star.notes).strip()
                
                branch_id = request.form.get('edit_branch_id')
                if branch_id: star.branch_id = branch_id
                
                is_active = request.form.get('edit_is_active')
                star.is_active = True if is_active == 'on' else False
                
                log_system_action("ĐỘI SAO ĐỎ", f"Đã cập nhật thông tin Sao đỏ {star.full_name}")
                flash(f"Đã cập nhật thông tin Sao đỏ {star.full_name}!", "success")
    except Exception as e:
        flash(f"Lỗi sửa thông tin sao đỏ: {e}", "error")
    return redirect(url_for('red_stars'))

@app.route('/delete_red_star/<int:id>', methods=['POST'])
def delete_red_star(id):
    try:
        with session_scope() as db_session:
            star = db_session.query(RedStar).filter(RedStar.id == id).first()
            if star:
                name = star.full_name
                # Xóa lịch trực liên quan
                db_session.query(Assignment).filter(Assignment.red_star_id == id).delete()
                
                # [BẢN VÁ LỖI]: Đổi 'red_star_id' thành 'evaluatee_id' cho khớp với CSDL
                db_session.query(StarEvaluation).filter(StarEvaluation.evaluatee_id == id).delete()
                
                # Xóa hồ sơ gốc
                db_session.delete(star)
                
                log_system_action("ĐỘI SAO ĐỎ", f"Đã xóa vĩnh viễn Sao đỏ {name} và các lịch trực liên quan")
                flash(f"Đã xóa vĩnh viễn Sao đỏ {name} và các lịch trực liên quan!", "success")
    except Exception as e:
        flash(f"Lỗi xóa sao đỏ: {e}", "error")
    return redirect(url_for('red_stars'))

@app.route('/import_red_stars', methods=['POST'])
def import_red_stars():
    if 'excel_file' not in request.files:
        flash("Không tìm thấy file tải lên!", "error")
        return redirect(url_for('red_stars'))
    
    file = request.files['excel_file']
    if file.filename == '':
        flash("Chưa chọn file nào!", "error")
        return redirect(url_for('red_stars'))
        
    if file and (file.filename.endswith('.xlsx') or file.filename.endswith('.xls')):
        try:
            df = pd.read_excel(file)
            df.columns = df.columns.str.strip().str.upper()
            
            col_name, col_phone, col_branch, col_gender, col_notes = None, None, None, None, None
            for col in df.columns:
                if col in ["HỌ VÀ TÊN", "HỌ TÊN", "TÊN"]: col_name = col
                if col in ["GIỚI TÍNH", "NAM/NỮ", "PHÁI"]: col_gender = col
                if col in ["LỚP", "CHI ĐOÀN"]: col_branch = col
                if col in ["SỐ ĐIỆN THOẠI", "SDT", "PHONE", "ĐIỆN THOẠI"]: col_phone = col
                if col in ["GHI CHÚ", "NOTE", "GHI CHU"]: col_notes = col
                    
            if not col_name or not col_branch:
                flash("Lỗi cấu trúc: File Excel bắt buộc phải có cột 'Họ và tên' và 'Chi đoàn/Lớp'.", "error")
                return redirect(url_for('red_stars'))
                
            with session_scope() as db_session:
                active_year = db_session.query(SchoolYear).filter(SchoolYear.is_active == True).first()
                if not active_year:
                    flash("Chưa có năm học nào được kích hoạt!", "error")
                    return redirect(url_for('red_stars'))
                    
                count_new, count_exist = 0, 0
                for _, row in df.iterrows():
                    if pd.isna(row[col_name]) or pd.isna(row[col_branch]): continue
                        
                    fullname = str(row[col_name]).strip()
                    branch_name = str(row[col_branch]).strip().upper()
                    phone_num = str(row[col_phone]).strip() if col_phone and not pd.isna(row[col_phone]) else ""
                    gender_val = str(row[col_gender]).strip().capitalize() if col_gender and not pd.isna(row[col_gender]) else "Nam"
                    notes_raw = row[col_notes] if col_notes else None
                    notes_val = str(notes_raw).strip() if pd.notna(notes_raw) and str(notes_raw).strip().lower() != 'nan' else ""
                    
                    branch = db_session.query(Branch).filter(Branch.name == branch_name, Branch.school_year_id == active_year.id).first()
                    if branch:
                        exist_rs = db_session.query(RedStar).filter(RedStar.full_name == fullname, RedStar.branch_id == branch.id).first()
                        if not exist_rs:
                            new_rs = RedStar(
                                full_name=fullname, gender=gender_val, phone=phone_num, 
                                branch_id=branch.id, notes=notes_val, is_active=True
                            )
                            db_session.add(new_rs)
                            count_new += 1
                        else: count_exist += 1
                            
                log_system_action("ĐỘI SAO ĐỎ", f"Đã nạp thành công {count_new} Sao đỏ mới từ Excel")
                flash(f"Đã nạp thành công {count_new} Sao đỏ mới (Bỏ qua {count_exist} học sinh bị trùng)", "success")
        except Exception as e:
            flash(f"Lỗi đọc file Excel: {str(e)}", "error")
    else:
        flash("Vui lòng chọn file Excel hợp lệ (.xlsx, .xls)", "error")
        
    return redirect(url_for('red_stars'))
# --- API LẤY LỊCH SỬ TRỰC CỦA TỪNG SAO ĐỎ ---
@app.route('/api/red_star_history/<int:star_id>')
def api_red_star_history(star_id):
    try:
        from database.database import session_scope
        from database.models import Assignment, DutyArea, RedStar
        import json
        import os
        
        with session_scope() as db_session:
            # 1. Kiểm tra sao đỏ
            star = db_session.query(RedStar).filter_by(id=star_id).first()
            if not star:
                return {"error": "Không tìm thấy thông tin Sao đỏ này!"}
                
            # 2. Đọc Sơ đồ Cụm trực để dịch từ Tên Khu Vực ra Tên Lớp
            zones_map = {}
            if os.path.exists("config/class_zones.json"):
                with open("config/class_zones.json", "r", encoding="utf-8") as f:
                    zones_map = json.load(f)
                    
            # 3. Quét toàn bộ lịch sử phân công
            history_data = []
            assignments = db_session.query(Assignment).filter(
                Assignment.red_star_id == star_id
            ).order_by(Assignment.week_number.desc(), Assignment.shift).all()
            
            for a in assignments:
                # Lấy tên khu vực
                area_name = a.duty_area.name if a.duty_area else "Không rõ"
                
                # Dịch ra các lớp phải trực
                classes = zones_map.get(area_name, [])
                class_str = ", ".join(classes) if classes else "Khu vực chung / Cổng"
                
                history_data.append({
                    "week": f"Tuần {a.week_number}",
                    "shift": a.shift,
                    "area": area_name,
                    "classes": class_str
                })
                
            return {"success": True, "data": history_data}
    except Exception as e:
        return {"error": str(e)}
    
# --- 1. QUẢN LÝ KHU VỰC TRỰC ---
@app.route('/duty-areas', methods=['GET', 'POST'])
def duty_areas():
    try:
        with session_scope() as db_session:
            active_year = db_session.query(SchoolYear).filter(SchoolYear.is_active == True).first()
            
            if request.method == 'POST':
                name = request.form.get('name', '').strip()
                try: req_stars = int(request.form.get('required_stars', 2))
                except: req_stars = 2
                
                selected_classes = request.form.getlist('branch_names')
                
                if name:
                    exist = db_session.query(DutyArea).filter_by(name=name).first()
                    if not exist:
                        new_area = DutyArea(name=name, required_stars=req_stars)
                        db_session.add(new_area)
                    else:
                        exist.required_stars = req_stars
                    
                    if selected_classes:
                        zones_map = {}
                        if os.path.exists("config/class_zones.json"):
                            with open("config/class_zones.json", "r", encoding="utf-8") as f:
                                zones_map = json.load(f)
                        
                        zones_map[name] = selected_classes
                        os.makedirs("config", exist_ok=True)
                        with open("config/class_zones.json", "w", encoding="utf-8") as f:
                            json.dump(zones_map, f, ensure_ascii=False, indent=4)
                            
                    log_system_action("CỤM TRỰC", f"Đã cập nhật thành công Khu vực/Cụm trực: {name}")
                    flash(f"Đã cập nhật thành công Khu vực/Cụm trực: {name}", "success")
                else:
                    flash("Vui lòng nhập tên khu vực!", "error")
                return redirect(url_for('duty_areas'))
                
            areas = db_session.query(DutyArea).all()
            branches = []
            if active_year:
                branches = db_session.query(Branch).filter(Branch.school_year_id == active_year.id).order_by(Branch.name).all()
            
            zones_map = {}
            assigned_classes = set() 
            if os.path.exists("config/class_zones.json"):
                with open("config/class_zones.json", "r", encoding="utf-8") as f:
                    try:
                        raw_map = json.load(f)
                    except:
                        raw_map = {}
                
                # [BẢN VÁ LỖI]: Tự động dọn rác - Chỉ giữ lại các cụm CÒN TỒN TẠI trong Database
                valid_area_names = [a.name for a in areas]
                zones_map = {k: v for k, v in raw_map.items() if k in valid_area_names}
                
                # Ghi đè lại file JSON cho sạch sẽ
                with open("config/class_zones.json", "w", encoding="utf-8") as f:
                    json.dump(zones_map, f, ensure_ascii=False, indent=4)
                    
                # Nạp lại danh sách lớp đã khóa
                for classes_in_zone in zones_map.values():
                    assigned_classes.update(classes_in_zone)
            
            # Lọc danh sách các lớp chưa được phân cụm
            unassigned_branches = [b for b in branches if b.name not in assigned_classes]
            
            return render_template('duty_areas.html', 
                                   areas=areas, 
                                   zones_map=zones_map, 
                                   branches=branches, 
                                   assigned_classes=assigned_classes,
                                   unassigned_branches=unassigned_branches) # Biến mới được truyền ra giao diện
    except Exception as e:
        flash(f"Lỗi phân hệ khu vực trực: {e}", "error")
        return redirect(url_for('dashboard'))

@app.route('/delete_duty_area/<int:id>', methods=['POST'])
def delete_duty_area(id):
    try:
        with session_scope() as db_session:
            area = db_session.query(DutyArea).filter(DutyArea.id == id).first()
            if area:
                zone_name = area.name
                
                # [VÁ LỖI TRIỆT ĐỂ]: Duyệt qua các phân công và xóa thủ công dựa trên ID object liên kết
                all_assignments = db_session.query(Assignment).all()
                for assign in all_assignments:
                    if assign.duty_area and assign.duty_area.id == id:
                        db_session.delete(assign)
                        
                db_session.delete(area)
                
                import os, json
                if os.path.exists("config/class_zones.json"):
                    with open("config/class_zones.json", "r", encoding="utf-8") as f:
                        zones_map = json.load(f)
                    if zone_name in zones_map:
                        del zones_map[zone_name]
                        with open("config/class_zones.json", "w", encoding="utf-8") as f:
                            json.dump(zones_map, f, ensure_ascii=False, indent=4)
                            
                log_system_action("CỤM TRỰC", f"Đã xóa khu vực trực: {zone_name}")
                flash("Đã xóa khu vực trực và các lịch trực liên quan thành công!", "success")
    except Exception as e:
        flash(f"Lỗi xóa khu vực: {e}", "error")
    return redirect(url_for('duty_areas'))
@app.route('/delete_multiple_duty_areas', methods=['POST'])
def delete_multiple_duty_areas():
    try:
        area_ids = request.form.getlist('area_ids')
        if not area_ids:
            flash("Vui lòng chọn ít nhất một cụm để xóa!", "warning")
            return redirect(url_for('duty_areas'))
        
        with session_scope() as db_session:
            # 1. Đọc file JSON 1 lần duy nhất trước khi vòng lặp chạy
            import os, json
            config_path = "config/class_zones.json"
            zones_map = {}
            if os.path.exists(config_path):
                with open(config_path, "r", encoding="utf-8") as f:
                    try: zones_map = json.load(f)
                    except: pass

            deleted_names = []
            
            # 2. Xóa từng cụm trong danh sách được chọn
            for a_id in area_ids:
                area = db_session.query(DutyArea).filter(DutyArea.id == int(a_id)).first()
                if area:
                    zone_name = area.name
                    deleted_names.append(zone_name)
                    
                    # Giải phóng lịch trực
                    all_assignments = db_session.query(Assignment).all()
                    for assign in all_assignments:
                        if assign.duty_area and assign.duty_area.id == area.id:
                            db_session.delete(assign)
                            
                    # Xóa cụm khỏi DB
                    db_session.delete(area)
                    
                    # Xóa cụm khỏi biến JSON
                    if zone_name in zones_map:
                        del zones_map[zone_name]
            
            # 3. Ghi lại file JSON 1 lần duy nhất để giải phóng các lớp
            os.makedirs("config", exist_ok=True)
            with open(config_path, "w", encoding="utf-8") as f:
                json.dump(zones_map, f, ensure_ascii=False, indent=4)
                
            log_system_action("CỤM TRỰC", f"Đã xóa nhiều cụm trực: {', '.join(deleted_names)}")
            flash(f"Đã xóa {len(deleted_names)} khu vực trực và giải phóng các lớp thành công!", "success")
    except Exception as e:
        flash(f"Lỗi xóa nhiều khu vực: {e}", "error")
        
    return redirect(url_for('duty_areas'))
@app.route('/edit_duty_area/<int:id>', methods=['POST'])
def edit_duty_area(id):
    try:
        with session_scope() as db_session:
            area = db_session.query(DutyArea).filter(DutyArea.id == id).first()
            if area:
                old_name = area.name
                new_name = request.form.get('edit_name', old_name).strip()
                try: new_req_stars = int(request.form.get('edit_required_stars', area.required_stars))
                except: new_req_stars = 2
                
                selected_classes = request.form.getlist('edit_branch_names')
                
                # Kiểm tra nếu đổi tên mà tên mới bị trùng
                if new_name != old_name:
                    exist = db_session.query(DutyArea).filter_by(name=new_name).first()
                    if exist:
                        flash(f"Tên khu vực '{new_name}' đã tồn tại!", "error")
                        return redirect(url_for('duty_areas'))
                        
                # Cập nhật Database
                area.name = new_name
                area.required_stars = new_req_stars
                
                # Cập nhật Sơ đồ lớp (File JSON)
                import os, json
                config_path = "config/class_zones.json"
                zones_map = {}
                if os.path.exists(config_path):
                    with open(config_path, "r", encoding="utf-8") as f:
                        try: zones_map = json.load(f)
                        except: pass
                        
                # Nếu đổi tên Cụm, xóa tên cũ trong JSON
                if old_name != new_name and old_name in zones_map:
                    del zones_map[old_name]
                    
                # Cập nhật danh sách lớp mới
                zones_map[new_name] = selected_classes
                
                os.makedirs("config", exist_ok=True)
                with open(config_path, "w", encoding="utf-8") as f:
                    json.dump(zones_map, f, ensure_ascii=False, indent=4)
                    
                log_system_action("CỤM TRỰC", f"Đã cập nhật khu vực: {old_name} -> {new_name}")
                flash(f"Đã cập nhật thông tin Cụm trực {new_name} thành công!", "success")
    except Exception as e:
        flash(f"Lỗi cập nhật khu vực trực: {e}", "error")
    return redirect(url_for('duty_areas'))

@app.route('/import_class_zones', methods=['POST'])
def import_class_zones():
    if 'excel_file' not in request.files:
        flash("Không tìm thấy file tải lên!", "error")
        return redirect(url_for('duty_areas'))
    
    file = request.files['excel_file']
    if file.filename == '':
        flash("Chưa chọn file nào!", "error")
        return redirect(url_for('duty_areas'))
        
    if file and (file.filename.endswith('.xlsx') or file.filename.endswith('.xls')):
        try:
            import numpy as np
            import pandas as pd
            import os
            import json
            
            df = pd.read_excel(file)
            df.columns = df.columns.str.strip().str.upper()
            
            if 'LỚP' not in df.columns:
                flash("Lỗi cấu trúc: File Excel bắt buộc phải có cột tiêu đề là 'LỚP'.", "error")
                return redirect(url_for('duty_areas'))
                
            dynamic_zones = {}
            cum_col = None
            for col in df.columns:
                if "CỤM" in col or "CUM" in col:
                    cum_col = col
                    break
            
            if cum_col:
                df[cum_col] = df[cum_col].replace('', np.nan).ffill()
                for index, row in df.iterrows():
                    zone_name = str(row[cum_col]).strip()
                    class_name = str(row['LỚP']).strip()
                    
                    if pd.notna(row['LỚP']) and class_name.lower() != 'nan' and class_name != '':
                        if zone_name.lower() != 'nan' and zone_name != '':
                            if zone_name not in dynamic_zones: 
                                dynamic_zones[zone_name] = []
                            dynamic_zones[zone_name].append(class_name)
            else:
                class_list = df['LỚP'].dropna().astype(str).tolist()
                grades = {}
                for c in class_list:
                    c_clean = c.strip()
                    if not c_clean or c_clean.lower() == 'nan': continue
                    grade = c_clean[:2]
                    if grade not in grades: grades[grade] = []
                    grades[grade].append(c_clean)
                for grade, classes in grades.items():
                    for i in range(0, len(classes), 2):
                        zone_name = f"CỤM_{grade}_{(i//2)+1:02d}"
                        dynamic_zones[zone_name] = classes[i:i+2]
            
            os.makedirs("config", exist_ok=True)
            with open("config/class_zones.json", "w", encoding="utf-8") as f:
                json.dump(dynamic_zones, f, ensure_ascii=False, indent=4)
                
            with session_scope() as db_session:
                count_new = 0
                for zone_name in dynamic_zones.keys():
                    exist = db_session.query(DutyArea).filter_by(name=zone_name).first()
                    if not exist:
                        # [ĐÃ ĐỒNG BỘ]: Set required_stars = 1 để khớp với luật "1 người/1 vị trí"
                        new_area = DutyArea(name=zone_name, required_stars=1)
                        db_session.add(new_area)
                        count_new += 1
                
            log_system_action("CỤM TRỰC", f"Đã nhập Sơ đồ lớp từ Excel, tạo {count_new} khu vực")
            flash(f"✅ Đã xử lý thành công file Sơ đồ lớp! Tạo {len(dynamic_zones)} Cụm trực và thêm mới {count_new} khu vực vào hệ thống.", "success")
        except Exception as e:
            flash(f"Lỗi đọc file Excel: {str(e)}", "error")
    else:
        flash("Vui lòng chọn file Excel hợp lệ (.xlsx, .xls)", "error")
    return redirect(url_for('duty_areas'))

# --- 2. PHÂN CÔNG LỊCH TRỰC ---
@app.route('/assignments', methods=['GET', 'POST'])
def assignments():
    try:
        with session_scope() as db_session:
            active_year = db_session.query(SchoolYear).filter(SchoolYear.is_active == True).first()
            
            # --- [NÂNG CẤP LÕI]: TỰ ĐỘNG NHẬN DIỆN TUẦN TRỰC MỚI NHẤT ---
            week_param = request.args.get('week')
            if week_param:
                current_week = int(week_param)
            else:
                # Quét CSDL tìm tuần mới nhất vừa được xếp lịch
                latest_assign = db_session.query(Assignment).order_by(Assignment.week_number.desc()).first()
                current_week = latest_assign.week_number if latest_assign else 1
            # -------------------------------------------------------------
            
            areas = db_session.query(DutyArea).all()
            active_stars = db_session.query(RedStar).filter(RedStar.is_active == True).all() 
            
            assign_list = []
            start_date_str = "" # Biến chứa ngày áp dụng để hiển thị lên Form
            
            if active_year:
                assign_list = db_session.query(Assignment).join(RedStar).join(Branch).filter(
                    Assignment.week_number == current_week,
                    Branch.school_year_id == active_year.id
                ).order_by(Assignment.shift, Assignment.id).all()
                
                # Trích xuất "Ngày áp dụng" từ bản ghi đầu tiên của Tuần này
                if assign_list and assign_list[0].date:
                    start_date_str = assign_list[0].date.strftime('%Y-%m-%d')
            
            return render_template('assignments.html', 
                                   active_year=active_year, 
                                   areas=areas, 
                                   stars=active_stars, 
                                   assign_list=assign_list, 
                                   current_week=current_week,
                                   start_date=start_date_str) # Truyền ngày ra màn hình HTML
    except Exception as e:
        flash(f"Lỗi tải lịch phân công: {e}", "error")
        return redirect(url_for('dashboard'))

@app.route('/delete_assignment/<int:id>', methods=['POST'])
def delete_assignment(id):
    try:
        with session_scope() as db_session:
            assign = db_session.query(Assignment).filter(Assignment.id == id).first()
            week = assign.week_number if assign else 1
            if assign:
                db_session.delete(assign)
                log_system_action("LỊCH TRỰC", f"Đã thu hồi 1 phân công trong Tuần {week}")
                flash("Đã thu hồi lịch phân công!", "success")
        return redirect(url_for('assignments', week=week))
    except Exception as e:
        flash(f"Lỗi thu hồi lịch: {e}", "error")
        return redirect(url_for('assignments'))

@app.route('/auto_assign', methods=['POST'])
def auto_assign():
    week_number = int(request.form.get('auto_week_number', 1))
    date_str = request.form.get('auto_date')
    # [BẢN VÁ MÚI GIỜ]: Ép giờ VN khi xếp lịch
    from datetime import datetime, timezone, timedelta
    vn_tz = timezone(timedelta(hours=7))

    try: start_date = datetime.strptime(date_str, '%Y-%m-%d').date() if date_str else datetime.now(vn_tz).date()
    except: start_date = datetime.now(vn_tz).date()
    
    shifts = request.form.getlist('shifts')
    if not shifts:
        flash("Vui lòng chọn ít nhất 1 ca trực!", "error")
        return redirect(url_for('assignments', week=week_number))
        
    try:
        with session_scope() as db_session:
            import random
            import json
            import os
            
            
            # 1. Xóa toàn bộ lịch cũ của tuần này để xếp lại từ đầu
            db_session.query(Assignment).filter(Assignment.week_number == week_number).delete()
            
            # 2. Lấy dữ liệu Cụm trực và Sao đỏ đang hoạt động
            areas = db_session.query(DutyArea).all()
            stars = db_session.query(RedStar).filter_by(is_active=True).all()
            
            if not areas or not stars:
                flash("Lỗi: Thiếu dữ liệu Khu vực trực hoặc Đội Sao đỏ để phân công!", "error")
                return redirect(url_for('assignments', week=week_number))
                
            # 3. Đọc cấu hình Sơ đồ lớp để né
            base_dir = os.path.dirname(os.path.abspath(__file__))
            config_path = os.path.join(base_dir, "config", "class_zones.json")
            zones_map = {}
            if os.path.exists(config_path):
                with open(config_path, "r", encoding="utf-8") as f:
                    try: zones_map = json.load(f)
                    except: pass
            
            def get_grade(class_name):
                match = re.search(r'(10|11|12)', str(class_name))
                return match.group(1) if match else ""

            # 4. TRÍCH XUẤT LỊCH SỬ XOAY VÒNG
            history_counts = {star.id: {} for star in stars}
            past_assignments = db_session.query(Assignment).filter(Assignment.week_number < week_number).all()
            
            for pa in past_assignments:
                if pa.red_star_id in history_counts and pa.duty_area:
                    area_id_val = pa.duty_area.id
                    history_counts[pa.red_star_id][area_id_val] = history_counts[pa.red_star_id].get(area_id_val, 0) + 1

            current_week_shift_counts = {star.id: 0 for star in stars}
            success_count = 0
            
            # QUY TẮC SẮP XẾP KHU VỰC: Ưu tiên Giám sát khối trước -> Cụm lớp -> Cổng sau cùng
            areas_sorted = sorted(areas, key=lambda a: 0 if "KHỐI" in a.name.upper() else (1 if "CỔNG" not in a.name.upper() else 2))
            
            # 5. Bắt đầu phân công cho từng ca (Sáng / Chiều)
            for shift in shifts:
                # Danh sách quân số rảnh trong ca này (Mỗi em chỉ dùng đúng 1 lần)
                available_stars = list(stars)
                random.shuffle(available_stars)
                
                for area in areas_sorted:
                    req_count = int(area.required_stars or 1)
                    assigned_count = 0
                    
                    area_name_lower = area.name.lower()
                    is_gate_area = "cổng" in area_name_lower
                    
                    area_classes = [c.strip().upper() for c in zones_map.get(area.name, [])]
                    area_grades = {get_grade(c) for c in area_classes if get_grade(c)}
                    
                    while assigned_count < req_count:
                        # Nếu vì lý do nào đó dùng hết quân mà vẫn thiếu ghế, cấp cứu nạp lại từ đầu
                        if not available_stars:
                            available_stars = list(stars)
                            
                        # Lọc danh sách thỏa mãn điều kiện an toàn
                        valid_stars = []
                        for star in available_stars:
                            star_class = star.branch.name.strip().upper() if star.branch else ""
                            star_grade = get_grade(star_class)
                            
                            is_conflict = False
                            
                            # Nếu không phải khu vực Cổng: Áp dụng luật nghiêm ngặt (Né khối + Né vị trí cũ)
                            if not is_gate_area:
                                if star_class in area_classes:
                                    is_conflict = True
                                elif star_grade and star_grade in area_grades:
                                    is_conflict = True
                                elif "KHỐI" in area.name.upper() and star_grade and star_grade in area.name:
                                    is_conflict = True
                                    
                                if history_counts[star.id].get(area.id, 0) > 0:
                                    is_conflict = True
                                    
                            if not is_conflict:
                                valid_stars.append(star)
                                
                        # Cứu hộ tầng 1: Nếu lọc quá ngặt mà rỗng, cho phép trực lại vị trí cũ nhưng vẫn né khối
                        if not valid_stars and not is_gate_area:
                            for star in available_stars:
                                star_class = star.branch.name.strip().upper() if star.branch else ""
                                star_grade = get_grade(star_class)
                                if star_class not in area_classes and (not star_grade or star_grade not in area_grades):
                                    valid_stars.append(star)
                                    
                        # Cứu hộ tầng 2 (Tuyệt đối): Lấy bất kỳ ai còn lại trong danh sách rảnh
                        if not valid_stars:
                            valid_stars = list(available_stars)
                            
                        # Sắp xếp theo ưu tiên: Ít lịch sử trực cụm này nhất & Ít việc trong tuần nhất
                        valid_stars.sort(key=lambda s: (
                            history_counts[s.id].get(area.id, 0),
                            current_week_shift_counts[s.id]
                        ))
                        
                        # CHỐT: Lấy em tốt nhất
                        chosen_star = valid_stars[0]
                        
                        new_assign = Assignment(
                            week_number=week_number,
                            shift=shift,
                            date=start_date,
                            red_star=chosen_star,      
                            duty_area=area      
                        )
                        db_session.add(new_assign)
                        
                        # XÓA VĨNH VIỄN EM NÀY KHỎI DANH SÁCH RẢNH CỦA CA ĐÓ (Đảm bảo mỗi vị trí là duy nhất 1 người)
                        available_stars.remove(chosen_star)
                        
                        current_week_shift_counts[chosen_star.id] += 1
                        history_counts[chosen_star.id][area.id] = history_counts[chosen_star.id].get(area.id, 0) + 1
                        assigned_count += 1
                        success_count += 1
            
            db_session.commit()
            
            if success_count > 0:
                log_system_action("LỊCH TRỰC", f"Đã phân công tự động Tuần {week_number}")
                flash(f"✅ Đã phân công thành công! (Mỗi vị trí là duy nhất 1 Sao đỏ, lấp đầy {success_count} vị trí)", "success")
            else:
                flash("⚠️ Không thể phân công do thiếu dữ liệu Sao đỏ.", "warning")
                
    except Exception as e:
        import traceback
        traceback.print_exc()
        flash(f"Lỗi hệ thống khi phân công: {str(e)}", "error")
        
    return redirect(url_for('assignments', week=week_number))

@app.route('/export_schedule/<int:week>')
def export_schedule(week):
    try:
        with session_scope() as db_session:
            active_year = db_session.query(SchoolYear).filter(SchoolYear.is_active == True).first()
            if not active_year:
                flash("Chưa có năm học kích hoạt!", "error")
                return redirect(url_for('assignments', week=week))
                
            # [ĐÃ VÁ LỖI POSTGRESQL]: Thêm .join(DutyArea) để hệ thống nhận diện được bảng Khu vực
            assignments = db_session.query(Assignment).join(RedStar).join(Branch).join(DutyArea).filter(
                Assignment.week_number == week,
                Branch.school_year_id == active_year.id
            ).order_by(Assignment.shift, DutyArea.name).all()
            
            if not assignments:
                flash(f"Chưa có dữ liệu lịch trực tuần {week} để xuất!", "error")
                return redirect(url_for('assignments', week=week))

            import os, json
            zones_map = {}
            if os.path.exists("config/class_zones.json"):
                with open("config/class_zones.json", "r", encoding="utf-8") as f:
                    zones_map = json.load(f)

            import openpyxl
            from openpyxl.styles import Font, Alignment, Border, Side
            import io
            from flask import send_file

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = f"Tuan_{week}"
            
            ws['A1'] = "ĐOÀN TRƯỜNG THPT THANH HÒA"
            ws['E1'] = "ĐOÀN TNCS HỒ CHÍ MINH"
            ws['A1'].font = Font(name="Times New Roman", size=11, bold=True)
            ws['E1'].font = Font(name="Times New Roman", size=11, bold=True)
            ws['E1'].alignment = Alignment(horizontal="right")
            
            ws['A3'] = f"LỊCH TRỰC ĐỘI SAO ĐỎ - TUẦN {week}"
            ws['A3'].font = Font(name="Times New Roman", size=14, bold=True)
            ws.merge_cells('A3:F3')
            ws['A3'].alignment = Alignment(horizontal="center")
            
            # [ĐÃ NÂNG CẤP LÕI]: Chuẩn hóa tiêu đề Excel tách bạch Cụm và Lớp
            headers = ["STT", "Họ Tên Sao Đỏ", "Lớp của SĐ", "Khu Vực Trực", "Giám Sát Các Lớp", "Ca Trực"]
            thin = Side(border_style="thin", color="000000")
            border = Border(left=thin, right=thin, top=thin, bottom=thin)
            
            for col_num, h_title in enumerate(headers, 1):
                c = ws.cell(row=5, column=col_num, value=h_title)
                c.font = Font(name="Times New Roman", size=12, bold=True)
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.border = border
                
            for idx, assign in enumerate(assignments, 1):
                area_name = assign.duty_area.name if assign.duty_area else "Chưa phân công"
                classes_list = zones_map.get(area_name, [])
                class_str = ", ".join(classes_list) if classes_list else "Khu vực chung"
                
                star_name = assign.red_star.full_name if assign.red_star else "Khuyết"
                branch_name = assign.red_star.branch.name if assign.red_star and assign.red_star.branch else "---"
                
                row_idx = idx + 5
                row_data = [idx, star_name, branch_name, area_name, class_str, assign.shift]
                
                for col_num, val in enumerate(row_data, 1):
                    c = ws.cell(row=row_idx, column=col_num, value=val)
                    c.font = Font(name="Times New Roman", size=12)
                    c.border = border
                    if col_num in [1, 3, 6]: 
                        c.alignment = Alignment(horizontal="center")
                
            ws.column_dimensions['A'].width = 8
            ws.column_dimensions['B'].width = 25
            ws.column_dimensions['C'].width = 15
            ws.column_dimensions['D'].width = 22
            ws.column_dimensions['E'].width = 30
            ws.column_dimensions['F'].width = 15
            
            log_system_action("XUẤT EXCEL", f"Xuất lịch trực chuyên nghiệp Tuần {week}")
            out = io.BytesIO()
            wb.save(out)
            out.seek(0)
            return send_file(out, download_name=f"Lich_Truc_Tuan_{week}.xlsx", as_attachment=True)
    except Exception as e:
        import traceback; traceback.print_exc()
        flash(f"Lỗi xuất excel lịch trực: {e}", "error")
        return redirect(url_for('assignments', week=week))
    
@app.route('/api/get_branch_duty_star')
def api_get_branch_duty_star():
    try:
        week_name = request.args.get('week', '')
        branch_param = request.args.get('branch_name', '').strip()
        
        if not week_name or not branch_param:
            return {"success": False, "error": "Thiếu tham số"}
            
        import re
        week_num_match = re.search(r'\d+', week_name)
        week_num = int(week_num_match.group()) if week_num_match else 0
        
        from database.database import session_scope
        from database.models import Assignment, DutyArea, Branch
        import os, json
        
        with session_scope() as db_session:
            # 1. Xử lý tên lớp an toàn tuyệt đối
            real_branch_name = branch_param
            if branch_param.isdigit():
                b_obj = db_session.query(Branch).filter_by(id=int(branch_param)).first()
                if b_obj: real_branch_name = b_obj.name
                    
            search_name = re.sub(r'\(.*?\)', '', real_branch_name)
            search_name = re.sub(r'(CHI ĐOÀN|CHI DOAN|LỚP|LOP)', '', search_name, flags=re.IGNORECASE)
            search_name = re.sub(r'[^A-Za-z0-9]', '', search_name).upper()
            
            # 2. Đọc file JSON với đường dẫn tuyệt đối (Chống lỗi File Not Found)
            zones_map = {}
            # Dùng app.root_path để trỏ chính xác tuyệt đối vào thư mục gốc của Flask
            config_path = os.path.join(app.root_path, "config", "class_zones.json")
            if not os.path.exists(config_path):
                config_path = "config/class_zones.json" # Dự phòng
                
            json_loaded_ok = False
            if os.path.exists(config_path):
                with open(config_path, "r", encoding="utf-8") as f:
                    try: 
                        zones_map = json.load(f)
                        json_loaded_ok = True
                    except: pass
                    
            # 3. Quét tìm cụm trực
            target_area_names = []
            for area_name, classes in zones_map.items():
                if isinstance(classes, list):
                    for c in classes:
                        c_clean = re.sub(r'\(.*?\)', '', str(c))
                        c_clean = re.sub(r'(CHI ĐOÀN|CHI DOAN|LỚP|LOP)', '', c_clean, flags=re.IGNORECASE)
                        c_clean = re.sub(r'[^A-Za-z0-9]', '', c_clean).upper()
                        
                        if search_name == c_clean:
                            target_area_names.append(area_name)
                            break
                            
            if not target_area_names:
                # [MÁY QUÉT 1]: Nếu lỗi ở File cấu hình hoặc quên lưu lớp vào cụm
                status_file = "OK" if json_loaded_ok else "LỖI ĐỌC FILE"
                return {"success": True, "star_name": f"Chưa phân cụm (Tìm: {search_name} - File: {status_file})"}
                
            # 4. Tìm lịch trực
            assignments = db_session.query(Assignment).join(DutyArea).filter(
                Assignment.week_number == week_num,
                DutyArea.name.in_(target_area_names)
            ).all()
            
            if not assignments:
                # [MÁY QUÉT 2]: Nếu lớp đã có Cụm, nhưng tuần hiện tại thầy chưa bấm nút Phân công
                return {"success": True, "star_name": f"Trống lịch trực (Tuần {week_num})"}
                
            # 5. Đóng gói kết quả
            star_info_list = []
            for asm in assignments:
                if asm.red_star:
                    s_name = asm.red_star.full_name.replace("SĐ: ", "").strip()
                    shift = asm.shift or ""
                    star_info_list.append(f"{s_name} ({shift})")
                    
            result_str = " | ".join(star_info_list) if star_info_list else "Đã phân cụm nhưng khuyết người trực"
            return {"success": True, "star_name": result_str}
            
    except Exception as e:
        return {"success": False, "error": str(e)}
    
@app.route('/api/get_swap_candidates/<int:assign_id>')
def api_get_swap_candidates(assign_id):
    try:
        from database.database import session_scope
        from database.models import Assignment, DutyArea, RedStar
        import json, os, re
        
        with session_scope() as db_session:
            assign = db_session.query(Assignment).filter_by(id=assign_id).first()
            if not assign: return {"error": "Không tìm thấy lịch trực"}
            
            week_num = assign.week_number
            shift = assign.shift
            
            current_area_name = assign.duty_area.name if assign.duty_area else ""
            current_star = assign.red_star
            current_star_id = current_star.id if current_star else 0
            
            active_stars = db_session.query(RedStar).filter_by(is_active=True).all()
            shift_assignments = db_session.query(Assignment).filter_by(week_number=week_num, shift=shift).all()
            busy_map = {a.red_star_id: (a.id, a.duty_area.name if a.duty_area else "") for a in shift_assignments}
            
            zones_map = {}
            if os.path.exists("config/class_zones.json"):
                with open("config/class_zones.json", "r", encoding="utf-8") as f:
                    try: zones_map = json.load(f)
                    except: pass
                    
            def get_grade(class_name):
                match = re.search(r'(10|11|12)', str(class_name))
                return match.group(1) if match else ""

            # [THUẬT TOÁN ĐỔI NGƯỜI CHÉO THÔNG MINH]: Kiểm tra 1 Học sinh có hợp lệ trực 1 Khu vực hay không
            def can_assign(star_obj, target_area_name):
                if not star_obj: return False
                if "cổng" in target_area_name.lower(): return True # Cổng thì vô tư
                
                s_class = star_obj.branch.name.strip().upper() if star_obj.branch else ""
                s_grade = get_grade(s_class)
                
                a_classes = [c.strip().upper() for c in zones_map.get(target_area_name, [])]
                a_grades = {get_grade(c) for c in a_classes if get_grade(c)}
                
                if s_class in a_classes: return False # Trùng lớp
                if s_grade and s_grade in a_grades: return False # Trùng khối
                if "KHỐI" in target_area_name.upper() and s_grade and s_grade in target_area_name.upper(): return False
                return True
            
            free_list = []
            busy_list = []
            
            for star in active_stars:
                if star.id == current_star_id: continue
                
                # 1. Học sinh này có đủ điều kiện thế chỗ vào vị trí hiện tại không?
                if not can_assign(star, current_area_name):
                    continue
                    
                is_busy = star.id in busy_map
                target_assign_id = None
                target_area_name = ""
                
                if is_busy:
                    target_assign_id, target_area_name = busy_map[star.id]
                    # 2. Học sinh hiện tại có đủ điều kiện sang thế chỗ ngược lại cho học sinh kia không?
                    if not can_assign(current_star, target_area_name):
                        continue
                
                item = {
                    "star_id": star.id,
                    "star_name": f"{star.full_name} ({star.branch.name if star.branch else ''})",
                    "target_assign_id": target_assign_id,
                    "target_area_name": target_area_name
                }
                
                if is_busy: busy_list.append(item)
                else: free_list.append(item)
                
            return {"free": free_list[:15], "busy": busy_list}
    except Exception as e:
        import traceback; traceback.print_exc()
        return {"error": str(e)}

@app.route('/execute_swap', methods=['POST'])
def execute_swap():
    assign_id = int(request.form.get('assign_id'))
    new_star_id = int(request.form.get('new_star_id'))
    target_assign_id = request.form.get('target_assign_id')
    
    try:
        with session_scope() as db_session:
            assign = db_session.query(Assignment).filter_by(id=assign_id).first()
            if not assign:
                flash("Lỗi: Không tìm thấy phân công gốc!", "error")
                return redirect(url_for('assignments'))
                
            week = assign.week_number
            if target_assign_id and target_assign_id != "None":
                target_assign = db_session.query(Assignment).filter_by(id=int(target_assign_id)).first()
                if target_assign:
                    temp = assign.red_star_id
                    assign.red_star_id = target_assign.red_star_id
                    target_assign.red_star_id = temp
                    log_system_action("LỊCH TRỰC", f"Hoán đổi vị trí trực chéo thành công trong Tuần {week}")
                    flash("✅ Đã hoán đổi chéo vị trí trực thành công!", "success")
            else:
                assign.red_star_id = new_star_id
                log_system_action("LỊCH TRỰC", f"Đổi người rảnh vào ca trực Tuần {week}")
                flash("✅ Đã thay thế người rảnh vào ca trực thành công!", "success")
                
            return redirect(url_for('assignments', week=week))
    except Exception as e:
        flash(f"Lỗi thực hiện đổi người trực: {e}", "error")
        return redirect(url_for('assignments'))
        
# --- 3. ĐÁNH GIÁ VÀ CHẤM ĐIỂM SAO ĐỎ (KPI 360 ĐỘ) ---

@app.route('/api/submit_evaluation', methods=['POST'])
def api_submit_evaluation():
    """API Nhận dữ liệu đánh giá trắc nghiệm - Bản bọc thép chống sập HTML"""
    import traceback # Import trực tiếp bên trong để đảm bảo luôn có hàm in lỗi
    
    try:
        # 1. BẮT DỮ LIỆU ĐẦU VÀO CỰC KỲ CHẶT CHẼ
        # Sử dụng get_json(force=True) để ép Flask đọc JSON dù thiếu Header
        data = request.get_json(silent=True, force=True) 
        if not data:
            data = request.form.to_dict() # Dự phòng nếu gửi bằng form
            
        if not data:
            # Trả về Dictionaries thuần túy, Flask hiện đại tự convert thành JSON
            return {"success": False, "error": "Máy chủ không nhận được dữ liệu (Payload trống)!"}, 400

        evaluatee_id_raw = data.get('evaluatee_id')
        week_name = data.get('week_name')
        
        if not evaluatee_id_raw or not week_name:
            return {"success": False, "error": "Thiếu mã học sinh hoặc tên tuần!"}, 400

        try:
            evaluatee_id = int(evaluatee_id_raw)
        except ValueError:
            return {"success": False, "error": "Mã học sinh không đúng định dạng số!"}, 400
        
        score_gio_giac = int(data.get('score_gio_giac', 5))
        score_tac_phong = int(data.get('score_tac_phong', 5))
        score_thai_do = int(data.get('score_thai_do', 5))
        score_cong_tam = int(data.get('score_cong_tam', 5))
        comment = str(data.get('comment', '')).strip()

        evaluator_username = session.get('username', 'Unknown')
        evaluator_role = session.get('role', 'Giáo viên/Đoàn')

        # 2. XỬ LÝ CƠ SỞ DỮ LIỆU
        with session_scope() as db_session:
            # Kiểm tra Sao đỏ
            target_star = db_session.query(RedStar).filter_by(id=evaluatee_id).first()
            if not target_star:
                return {"success": False, "error": "Học sinh Sao đỏ này không còn tồn tại trong hệ thống!"}, 404

            # Cập nhật hoặc tạo mới
            existing_eval = db_session.query(StarEvaluation).filter_by(
                evaluator_username=str(evaluator_username),
                evaluatee_id=evaluatee_id,
                week_name=str(week_name)
            ).first()
            
            if existing_eval:
                existing_eval.score_gio_giac = score_gio_giac
                existing_eval.score_tac_phong = score_tac_phong
                existing_eval.score_thai_do = score_thai_do
                existing_eval.score_cong_tam = score_cong_tam
                existing_eval.comment = comment
                msg = "Đã lưu bản cập nhật phiếu đánh giá thành công!"
            else:
                new_eval = StarEvaluation(
                    evaluator_username=str(evaluator_username),
                    evaluator_role=str(evaluator_role),
                    evaluatee_id=evaluatee_id,
                    week_name=str(week_name),
                    score_gio_giac=score_gio_giac,
                    score_tac_phong=score_tac_phong,
                    score_thai_do=score_thai_do,
                    score_cong_tam=score_cong_tam,
                    comment=comment
                )
                db_session.add(new_eval)
                msg = "Đã nộp phiếu đánh giá mới thành công!"
                
            log_system_action("ĐÁNH GIÁ KPI", f"User {evaluator_username} chấm điểm Sao đỏ ID {evaluatee_id}")
            
            # Trả về thành công
            return {"success": True, "message": msg}, 200

    except Exception as e:
        traceback.print_exc() # In màn hình đen (Console)
        # Bắt mọi lỗi sập nguồn và trả về JSON an toàn
        return {"success": False, "error": f"Lỗi nội bộ Server: {str(e)}"}, 500
    
@app.route('/star-evaluations', methods=['GET', 'POST'])
def star_evaluations():
    """Trang quản trị xem và quản lý kết quả KPI (Dành cho Admin/Bí thư)"""
    try:
        with session_scope() as db_session:
            active_year = db_session.query(SchoolYear).filter(SchoolYear.is_active == True).first()
            
            # Tự động nhận diện Tuần Đánh giá mới nhất
            week_param = request.args.get('week')
            if week_param:
                current_week = int(week_param)
            else:
                latest_eval = db_session.query(StarEvaluation).order_by(StarEvaluation.id.desc()).first()
                if latest_eval:
                    # Lấy số từ chuỗi (Ví dụ "Tuần 5" -> 5)
                    
                    match = re.search(r'\d+', latest_eval.week_name)
                    current_week = int(match.group()) if match else 1
                else:
                    latest_assign = db_session.query(Assignment).order_by(Assignment.week_number.desc()).first()
                    current_week = latest_assign.week_number if latest_assign else 1

            week_name_str = f"Tuần {current_week}"
            stars = db_session.query(RedStar).filter(RedStar.is_active == True).all()
            
            # Lấy toàn bộ đánh giá của tuần
            evaluations = db_session.query(StarEvaluation).filter(StarEvaluation.week_name == week_name_str).all()
            
            # Gom nhóm đánh giá theo từng Sao đỏ để tính KPI trung bình
            eval_dict = {}
            for e in evaluations:
                if e.evaluatee_id not in eval_dict:
                    eval_dict[e.evaluatee_id] = []
                eval_dict[e.evaluatee_id].append(e)

            # Tính toán Tổng điểm KPI hiển thị ra màn hình
            kpi_summary = {}
            for star in stars:
                if star.id in eval_dict:
                    evals = eval_dict[star.id]
                    total_points = sum((ev.score_gio_giac + ev.score_tac_phong + ev.score_thai_do + ev.score_cong_tam) for ev in evals)
                    avg_score = total_points / len(evals) # Điểm trung bình / Lượt đánh giá
                    kpi_percent = (avg_score / 20) * 100  # Quy ra % (Tối đa 20đ/lượt)
                    
                    kpi_summary[star.id] = {
                        'count': len(evals),
                        'avg_score': round(avg_score, 1),
                        'kpi_percent': round(kpi_percent, 1),
                        'evals_detail': evals
                    }

            return render_template('star_evaluations.html', 
                                   active_year=active_year, 
                                   stars=stars, 
                                   kpi_summary=kpi_summary, 
                                   current_week=current_week)
    except Exception as e:
        import traceback; traceback.print_exc()
        flash(f"Lỗi phân hệ đánh giá KPI: {e}", "error")
        return redirect(url_for('dashboard'))

# ==========================================
# MODULE: XẾP HẠNG THI ĐUA ĐỘI SAO ĐỎ (THÁNG / HỌC KỲ / NĂM)
# ==========================================
@app.route('/star-ranking/<report_type>', methods=['GET', 'POST'])
def star_ranking(report_type):
    try:
        with session_scope() as db_session:
            active_year = db_session.query(SchoolYear).filter(SchoolYear.is_active == True).first()
            available_weeks = []
            if active_year:
                # Tìm các tuần đã có dữ liệu đánh giá
                weeks_db = db_session.query(StarEvaluation.week_name).distinct().all()
                
                week_nums = []
                for w in weeks_db:
                    match = re.search(r'\d+', w[0])
                    if match: week_nums.append(int(match.group()))
                available_weeks = sorted(list(set(week_nums)))

            ranking_data = []
            selected_title = ""
            selected_weeks = []
            display_selected_weeks = []

            if request.method == 'POST' or report_type == 'yearly':
                if report_type == 'monthly':
                    selected_title = request.form.get('title', 'Tháng 9')
                    display_selected_weeks = [int(w) for w in request.form.getlist('weeks')]
                    selected_weeks = [f"Tuần {w}" for w in display_selected_weeks]
                elif report_type == 'semester':
                    selected_title = request.form.get('title', 'Học Kỳ 1')
                    display_selected_weeks = [int(w) for w in request.form.getlist('weeks')]
                    selected_weeks = [f"Tuần {w}" for w in display_selected_weeks]
                elif report_type == 'yearly':
                    selected_title = f"Năm học {active_year.name}" if active_year else "Năm học"
                    display_selected_weeks = available_weeks
                    selected_weeks = [f"Tuần {w}" for w in display_selected_weeks]

                if not selected_weeks and report_type != 'yearly':
                    flash("Vui lòng chọn ít nhất 1 tuần để tổng hợp điểm!", "error")
                elif active_year:
                    stars = db_session.query(RedStar).filter(RedStar.is_active == True).all()
                    for star in stars:
                        if star.branch and star.branch.school_year_id == active_year.id:
                            # [NÂNG CẤP LÕI]: Dùng evaluatee_id theo cấu trúc mới
                            evals = db_session.query(StarEvaluation).filter(
                                StarEvaluation.evaluatee_id == star.id,
                                StarEvaluation.week_name.in_(selected_weeks)
                            ).all()
                            
                            if evals:
                                # [NÂNG CẤP LÕI]: Tính tổng từ 4 cột trắc nghiệm
                                total_score = sum((e.score_gio_giac + e.score_tac_phong + e.score_thai_do + e.score_cong_tam) for e in evals)
                                eval_count = len(evals)
                                avg_score = total_score / eval_count
                                kpi_percent = (avg_score / 20) * 100
                                
                                ranking_data.append({
                                    'star_name': star.full_name,
                                    'branch_name': star.branch.name,
                                    'total_score': total_score,
                                    'eval_count': eval_count,
                                    'avg_score': round(avg_score, 1),
                                    'kpi_percent': round(kpi_percent, 1)
                                })
                    
                    # Sắp xếp dựa trên KPI Trung bình để công bằng cho người trực ít / trực nhiều
                    ranking_data = sorted(ranking_data, key=lambda x: x['avg_score'], reverse=True)
                    current_rank = 1
                    for i, d in enumerate(ranking_data):
                        if i > 0 and d['avg_score'] < ranking_data[i-1]['avg_score']:
                            current_rank = i + 1
                        d['rank'] = current_rank

            if report_type == 'monthly': page_title = "Đánh giá Tháng"
            elif report_type == 'semester': page_title = "Đánh giá Học kỳ"
            else: page_title = "Tổng kết Năm học"
            
            return render_template(
                'star_ranking.html', 
                report_type=report_type,
                page_title=page_title,
                available_weeks=available_weeks, 
                ranking_data=ranking_data,
                selected_title=selected_title, 
                selected_weeks=display_selected_weeks,
                active_year=active_year
            )
    except Exception as e:
        import traceback; traceback.print_exc()
        flash(f"Lỗi hệ thống khi tổng hợp đánh giá Sao đỏ: {e}", "error")
        return redirect(url_for('dashboard'))

# ==========================================
# API: KIỂM TRA VÀ ĐỔI TRẠNG THÁI KHÓA SỔ TUẦN (BẢO TOÀN LỖI ĐÃ PHÚC KHẢO)
# ==========================================
@app.route('/api/toggle_week_lock', methods=['POST'])
def api_toggle_week_lock():
    try:
        data = request.get_json()
        week_name = data.get('week_name')
        
        with session_scope() as db_session:
            active_year = db_session.query(SchoolYear).filter_by(is_active=True).first()
            if not active_year: return {"success": False, "error": "Chưa có năm học kích hoạt!"}
                
            scores = db_session.query(WeeklyScore).join(Branch).filter(WeeklyScore.week == week_name, Branch.school_year_id == active_year.id).all()
            new_status = not any(getattr(s, 'is_locked', False) for s in scores)
            
            for s in scores:
                if new_status == True and not s.is_locked:
                    if s.note:
                        all_categories = db_session.query(ViolationCategory).filter_by(school_year_id=s.branch.school_year_id).all()
                        sorted_cats = sorted(all_categories, key=lambda x: len(x.name), reverse=True)
                        parsed_errors = {}
                        
                        parts = smart_split_note(s.note)
                        for part in parts:
                            part_clean = part.strip()
                            if not part_clean or "vắng 0" in part_clean.lower(): continue
                            
                            # ==========================================================
                            # [KIM BÀI MIỄN TỬ]: BỎ QUA HOÀN TOÀN CÁC LỖI ĐÃ ĐƯỢC GỠ
                            # ==========================================================
                            if "(đã gỡ)" in part_clean.lower():
                                parsed_errors[("ĐÃ_GỠ", part_clean.lower(), part_clean, "")] = 1
                                continue
                            
                            # 1. BÓC TÁCH THẺ NGÀY BẤT KỂ VỊ TRÍ
                            match_day = re.search(r'\[\s*(T[2-7]|CN)[^\]]*\]|\(\s*(T[2-7]|CN)[^\)]*\)', part_clean, re.IGNORECASE)
                            day_pfx = match_day.group(0).upper().replace('(', '[').replace(')', ']') if match_day else ""
                            if match_day: part_clean = part_clean.replace(match_day.group(0), "")
                            
                            # 2. BÓC TÁCH TÊN HỌC SINH (Lúc này thẻ ngày đã bị xóa, không thể nhầm lẫn)
                            match_stu = re.search(r'\[(.*?)\]|\((.*?)\)', part_clean)
                            stu_name_raw = ""
                            if match_stu:
                                stu_name_raw = match_stu.group(1) if match_stu.group(1) else match_stu.group(2)
                                part_clean = part_clean.replace(match_stu.group(0), "")
                            
                            matched = False
                            for cat in sorted_cats:
                                if cat.name.lower() in part_clean.lower():
                                    match_qty = re.search(r'(?:x|:|-)\s*(\d+)', part_clean.lower())
                                    qty = int(match_qty.group(1)) if match_qty else 1
                                    
                                    if not stu_name_raw: # Hỗ trợ quên ngoặc vuông
                                        temp = re.sub(re.escape(cat.name), '', part_clean, flags=re.IGNORECASE)
                                        temp = re.sub(r'(?:x|:|-)\s*\d+', '', temp, flags=re.IGNORECASE)
                                        stu_name_raw = temp.strip()
                                        
                                    stu_name_normalized = " ".join(str(stu_name_raw).split()).title() if stu_name_raw else ""
                                    key = (cat.name, stu_name_normalized.lower(), stu_name_normalized, day_pfx)
                                    parsed_errors[key] = parsed_errors.get(key, 0) + qty
                                    matched = True
                                    break
                            
                            if not matched:
                                parsed_errors[("MANUAL", part_clean.lower(), part_clean.strip(), day_pfx)] = 1
                                
                        final_parts = []
                        db_session.query(WeeklyViolation).filter_by(weekly_score_id=s.id).delete()
                        
                        for (cat_name, stu_key, stu_display, day_pfx), qty in parsed_errors.items():
                            if cat_name == "MANUAL":
                                final_parts.append(f"{day_pfx} {stu_display}".strip())
                            elif cat_name == "ĐÃ_GỠ":
                                final_parts.append(stu_display) # Bê nguyên xi chuỗi (ĐÃ GỠ) vào lại mà không trừ điểm!
                            else:
                                base_str = f"{cat_name} x{qty} [{stu_display}]" if stu_display else f"{cat_name} x{qty}"
                                final_parts.append(f"{day_pfx} {base_str}".strip())
                                
                                safe_stu = f"{stu_display} {day_pfx}".strip() if stu_display else day_pfx
                                cat_obj = next((c for c in all_categories if c.name == cat_name), None)
                                if cat_obj and getattr(cat_obj, 'point_type', 'Điểm trừ') != 'Điểm cộng':
                                    db_session.add(WeeklyViolation(weekly_score_id=s.id, violation_id=cat_obj.id, quantity=qty, student_name=safe_stu if safe_stu else None))
                                
                        s.note = " ; ".join(final_parts)
                
                if new_status == True and getattr(s, 'evidence_image', None):
                    s.evidence_image = None # Dọn rác ảnh

                s.is_locked = new_status
                
            status_text = "Khóa sổ (Đã chốt)" if new_status else "Mở khóa sổ"
            return {"success": True, "is_locked": new_status, "message": f"Đã {status_text} thành công {week_name}!"}
    except Exception as e:
        import traceback; traceback.print_exc()
        return {"success": False, "error": str(e)}

@app.route('/api/get_week_lock_status/<week_name>')
def api_get_week_lock_status(week_name):
    try:
        with session_scope() as db_session:
            active_year = db_session.query(SchoolYear).filter_by(is_active=True).first()
            if not active_year: return {"is_locked": False}
            
            scores = db_session.query(WeeklyScore).join(Branch).filter(
                WeeklyScore.week == week_name,
                Branch.school_year_id == active_year.id
            ).all()
            
            is_locked = any(getattr(s, 'is_locked', False) for s in scores)
            return {"is_locked": is_locked}
    except Exception as e:
        return {"is_locked": False}
# ==========================================
# HÀM HỖ TRỢ: BÓC TÁCH LỖI THÔNG MINH (CHỐNG CẮT NHẦM DẤU PHẨY)
# ==========================================
def smart_split_note(note_str):
    if not note_str: return []
    parts = []
    current_part = []
    in_bracket = 0
    for char in str(note_str):
        if char in '[(': in_bracket += 1
        elif char in '])': in_bracket -= 1
        
        # Chỉ cắt chuỗi khi gặp dấu phẩy/chấm phẩy và ĐANG KHÔNG NẰM TRONG NGOẶC
        if char in ',;+\n' and in_bracket <= 0:
            parts.append(''.join(current_part))
            current_part = []
        else:
            current_part.append(char)
    if current_part:
        parts.append(''.join(current_part))
    return [p.strip() for p in parts if p.strip()]
# ==========================================
# MODULE: Thêm hàm xử lý khử trùng lặp vào file
# ==========================================


def reconcile_same_day_absences(note_string):
    """
    Hàm đối chiếu và khử trùng lặp lỗi vắng có hỗ trợ khớp tên thông minh 
    (Xử lý trường hợp Sổ Đầu Bài ghi 'Nguyễn Văn An' còn Sao Đỏ chỉ ghi 'An').
    """
    if not note_string:
        return ""
        
    parts = [p.strip() for p in str(note_string).split(';') if p.strip()]
    seen_absences = set() 
    filtered_parts = []
    
    absence_keywords = ["vắng", "nghỉ", "không phép", "trốn tiết"]

    for part in parts:
        part_lower = part.lower()
        is_absence = any(kw in part_lower for kw in absence_keywords)
        
        if is_absence:
            # 1. Trích xuất tiền tố ngày (VD: [CN], [T2]...)
            day_match = re.search(r'\[(T[2-7](?:\s*Chiều|\s*Chieu)?|CN)\]', part, re.IGNORECASE)
            day_pfx = day_match.group(0).upper() if day_match else ""
            
            # 2. Trích xuất tên học sinh trong ngoặc
            all_brackets = re.findall(r'\[(.*?)\]', part)
            raw_student_name = ""
            for b_val in all_brackets:
                if b_val.upper() not in ["T2", "T3", "T4", "T5", "T6", "T7", "CN", "T2 CHIỀU", "T3 CHIỀU", "T4 CHIỀU", "T5 CHIỀU", "T6 CHIỀU", "T7 CHIỀU"]:
                    raw_student_name = b_val.strip()
                    break
            
            if raw_student_name and day_pfx:
                # [THUẬT TOÁN KHỚP TÊN MỜ]: Lấy từ cuối cùng (Tên chính) để so sánh bất chấp họ đệm dài ngắn
                name_tokens = raw_student_name.split()
                first_name_key = name_tokens[-1].lower() if name_tokens else raw_student_name.lower()
                
                absence_key = (first_name_key, day_pfx)
                if absence_key in seen_absences:
                    # Đã có bản ghi vắng trùng tên chính vào cùng ngày -> BỎ QUA để chống trừ 2 lần!
                    continue
                else:
                    seen_absences.add(absence_key)
        
        filtered_parts.append(part)
        
    return " ; ".join(filtered_parts)
# ==========================================
# MODULE: NHẬP ĐIỂM TUẦN & TỰ ĐỘNG BÓC TÁCH LỖI VÀO SỔ ĐEN (BẢO TOÀN LỖI ĐÃ PHÚC KHẢO)
# ==========================================
@app.route('/weekly', methods=['GET', 'POST'])
def weekly():
    try:
        with session_scope() as db_session:
            active_year = db_session.query(SchoolYear).filter(SchoolYear.is_active == True).first()
            
            if request.method == 'POST':
                if not active_year: flash("Chưa có năm học nào được kích hoạt!", "error"); return redirect(url_for('weekly'))
                    
                week_name = request.form.get('week_name', 'Tuần 1')
                scores_check = db_session.query(WeeklyScore).join(Branch).filter(WeeklyScore.week == week_name, Branch.school_year_id == active_year.id).all()
                if any(getattr(s, 'is_locked', False) for s in scores_check):
                    flash(f"⚠️ {week_name} đã được chốt sổ (khóa điểm)! Toàn bộ dữ liệu đã được đóng băng, không thể chỉnh sửa.", "error")
                    return redirect(url_for('weekly', week=week_name))

                start_date_val = request.form.get('start_date', '').strip()
                end_date_val = request.form.get('end_date', '').strip()
                branches = db_session.query(Branch).filter(Branch.school_year_id == active_year.id).all()
                
                DIEM_8, DIEM_9, DIEM_10 = 1.0, 3.0, 5.0; TUAN_KHA, TUAN_TOT = 20.0, 30.0
                settings = db_session.query(ScoreSettings).filter_by(school_year_id=active_year.id).first()
                if settings:
                    DIEM_8, DIEM_9, DIEM_10 = float(settings.diem_8), float(settings.diem_9), float(settings.diem_10)
                    TUAN_KHA, TUAN_TOT = float(settings.diem_tuan_kha), float(settings.diem_tuan_tot)
                    
                all_categories = db_session.query(ViolationCategory).filter_by(school_year_id=active_year.id).all() if active_year else []
                actor_username = session.get('username', 'Hệ thống'); actor_fullname = session.get('full_name', 'Người dùng')
                
                for branch in branches:
                    b_id = str(branch.id); b_group = str(branch.group) if branch.group else "1"
                    rating = request.form.get(f'rating_{b_id}', 'Bình thường')
                    try: c_8 = int(request.form.get(f'c8_{b_id}', 0) or 0)
                    except: c_8 = 0
                    try: c_9 = int(request.form.get(f'c9_{b_id}', 0) or 0)
                    except: c_9 = 0
                    try: c_10 = int(request.form.get(f'c10_{b_id}', 0) or 0)
                    except: c_10 = 0
                    try: truc = float(request.form.get(f'truc_{b_id}', 100.0) or 100.0)
                    except: truc = 100.0
                    try: cong = float(request.form.get(f'cong_{b_id}', 0.0) or 0.0)
                    except: cong = 0.0
                    
                    note = request.form.get(f'note_{b_id}', '').strip()
                    note = reconcile_same_day_absences(note)
                    
                    max_tot_web = int(getattr(settings, 'max_diem_tot', 14)) if settings else 14
                    tong_sl_diem = c_8 + c_9 + c_10
                    if tong_sl_diem > max_tot_web:
                        lech = tong_sl_diem - max_tot_web
                        x8 = min(c_8, lech); c_8 -= x8; lech -= x8
                        x9 = min(c_9, lech); c_9 -= x9; lech -= x9
                        c_10 -= lech
                    
                    diem_quy_uoc = 0.0
                    if "1" in b_group: diem_quy_uoc = (c_9 * DIEM_9) + (c_10 * DIEM_10)
                    elif "2" in b_group: diem_quy_uoc = (c_8 * DIEM_8) + (c_9 * DIEM_9) + (c_10 * DIEM_10)
                    else: diem_quy_uoc = (c_9 * DIEM_9) + (c_10 * DIEM_10)

                    diem_xep_loai = 0.0
                    if rating == "Tuần Tốt": diem_xep_loai = TUAN_TOT
                    elif rating == "Tuần Khá": diem_xep_loai = TUAN_KHA

                    diem_tru_auto = 0.0; new_violations = [] 
                    
                    if note:
                        parts = smart_split_note(note)
                        sorted_cats = sorted(all_categories, key=lambda x: len(x.name), reverse=True)
                        parsed_errors = {}
                        
                        for part in parts:
                            part_clean = part.strip()
                            if not part_clean: continue
                            if "Vắng: 0" in part_clean.lower() or "vắng 0" in part_clean.lower(): continue
                            
                            # ==========================================================
                            # [KIM BÀI MIỄN TỬ]: BỎ QUA HOÀN TOÀN CÁC LỖI ĐÃ ĐƯỢC GỠ
                            # ==========================================================
                            if "(đã gỡ)" in part_clean.lower():
                                parsed_errors[("ĐÃ_GỠ", part_clean.lower(), part_clean, "")] = 1
                                continue
                            
                            # 1. BÓC TÁCH THẺ NGÀY (Cất đi)
                            match_day = re.search(r'\[\s*(T[2-7]|CN)[^\]]*\]|\(\s*(T[2-7]|CN)[^\)]*\)', part_clean, re.IGNORECASE)
                            day_pfx = match_day.group(0).upper().replace('(', '[').replace(')', ']') if match_day else ""
                            
                            # 2. Xóa thẻ ngày khỏi câu để không nhiễu
                            if match_day: part_clean = part_clean.replace(match_day.group(0), "")
                            
                            # 3. TÌM TÊN HỌC SINH TỪ PHẦN CÒN LẠI
                            match_stu = re.search(r'\[(.*?)\]|\((.*?)\)', part_clean)
                            stu_name_raw = ""
                            if match_stu:
                                stu_name_raw = match_stu.group(1) if match_stu.group(1) else match_stu.group(2)
                                part_clean = part_clean.replace(match_stu.group(0), "") # Xóa luôn tên HS để dễ tìm Lỗi
                            
                            matched = False
                            for cat in sorted_cats:
                                if cat.name.lower() in part_clean.lower():
                                    match_qty = re.search(r'(?:x|:|-)\s*(\d+)', part_clean.lower())
                                    qty = int(match_qty.group(1)) if match_qty else 1 
                                    
                                    if not stu_name_raw: # Hỗ trợ quên ngoặc vuông
                                        temp = re.sub(re.escape(cat.name), '', part_clean, flags=re.IGNORECASE)
                                        temp = re.sub(r'(?:x|:|-)\s*\d+', '', temp, flags=re.IGNORECASE)
                                        stu_name_raw = temp.strip()
                                        
                                    stu_name_normalized = " ".join(str(stu_name_raw).split()).title() if stu_name_raw else ""
                                    key = (cat.name, stu_name_normalized.lower(), stu_name_normalized, day_pfx)
                                    parsed_errors[key] = parsed_errors.get(key, 0) + qty
                                    matched = True
                                    break
                                    
                            if not matched:
                                parsed_errors[("MANUAL", part_clean.lower(), part_clean.strip(), day_pfx)] = 1

                        max_tot_bad = int(getattr(settings, 'max_diem_tot', 14)) if settings else 14
                        max_mon_bad = int(getattr(settings, 'max_diem_mon', 4)) if settings else 4
                        bad_marks_expanded = []
                        other_errors = []
                        
                        for (cat_name, stu_key, stu_display, day_pfx), qty in parsed_errors.items():
                            if cat_name == "MANUAL" or cat_name == "ĐÃ_GỠ":
                                other_errors.append((cat_name, stu_display, qty, day_pfx))
                            else:
                                is_bad_mark = "không học bài" in cat_name.lower() or "điểm kém" in cat_name.lower()
                                if is_bad_mark:
                                    mon_match = re.search(r'\(Môn (.*?)\)', stu_display, re.IGNORECASE) if stu_display else None
                                    mon = mon_match.group(1).strip() if mon_match else "Khác"
                                    for _ in range(qty): bad_marks_expanded.append({'cat_name': cat_name, 'stu_display': stu_display, 'mon': mon, 'day_pfx': day_pfx})
                                else:
                                    other_errors.append((cat_name, stu_display, qty, day_pfx))

                        bad_by_subj = {}
                        for bm in bad_marks_expanded: bad_by_subj.setdefault(bm['mon'], []).append(bm)
                            
                        surviving_bad_marks = []
                        for marks in bad_by_subj.values(): surviving_bad_marks.extend(marks[:max_mon_bad])
                        surviving_bad_marks = surviving_bad_marks[:max_tot_bad]

                        capped_bad_counts = {}
                        for bm in surviving_bad_marks:
                            k = (bm['cat_name'], bm['stu_display'], bm['day_pfx'])
                            capped_bad_counts[k] = capped_bad_counts.get(k, 0) + 1

                        final_note_parts = []
                        for (cat_name, stu_display, day_pfx), qty in capped_bad_counts.items():
                            base_str = f"{cat_name} x{qty} [{stu_display}]" if stu_display else f"{cat_name} x{qty}"
                            final_note_parts.append(f"{day_pfx} {base_str}".strip())
                            for cat in sorted_cats:
                                if cat.name == cat_name and getattr(cat, 'point_type', 'Điểm trừ') != 'Điểm cộng':
                                    diem_tru_auto += float(cat.penalty_points * qty)
                                    safe_stu = f"{stu_display} {day_pfx}".strip() if stu_display else day_pfx
                                    new_violations.append({'violation_id': cat.id, 'quantity': qty, 'student_name': safe_stu if safe_stu else None})
                                    break
                                    
                        for cat_name, stu_display, qty, day_pfx in other_errors:
                            if cat_name == "MANUAL":
                                final_note_parts.append(f"{day_pfx} {stu_display}".strip() if day_pfx else stu_display)
                            elif cat_name == "ĐÃ_GỠ":
                                final_note_parts.append(stu_display) # Bê nguyên xi chuỗi đã gỡ vào lại
                            else:
                                base_str = f"{cat_name} x{qty} [{stu_display}]" if stu_display else f"{cat_name} x{qty}"
                                final_note_parts.append(f"{day_pfx} {base_str}".strip())
                                for cat in sorted_cats:
                                    if cat.name == cat_name and getattr(cat, 'point_type', 'Điểm trừ') != 'Điểm cộng':
                                        diem_tru_auto += float(cat.penalty_points * qty)
                                        safe_stu = f"{stu_display} {day_pfx}".strip() if stu_display else day_pfx
                                        new_violations.append({'violation_id': cat.id, 'quantity': qty, 'student_name': safe_stu if safe_stu else None})
                                        break

                        note = " ; ".join(final_note_parts)
                        
                    tong_diem_tru = diem_tru_auto
                    tru = diem_tru_auto 
                    total_val = truc + diem_xep_loai + diem_quy_uoc + cong - tong_diem_tru
                    
                    score = db_session.query(WeeklyScore).filter_by(branch_id=branch.id, week=week_name).first()
                    old_val = float(score.total_score) if score and score.total_score is not None else None
                    
                    if score:
                        if old_val is not None and old_val != total_val:
                            log_details = f"Sửa điểm lớp {branch.name} ({week_name}): Từ {old_val}đ thành {total_val}đ"
                            db_session.add(ActionLog(username=actor_username, full_name=actor_fullname, action_type="THAY ĐỔI ĐIỂM", details=log_details))
                            
                        score.week_rating = rating; score.count_8 = c_8; score.count_9 = c_9; score.count_10 = c_10; score.score_truc = truc; score.score_cong = cong; score.score_tru = tru; score.note = note; score.total_score = total_val; score.start_date = start_date_val; score.end_date = end_date_val
                    else:
                        score = WeeklyScore(branch_id=branch.id, week=week_name, week_rating=rating, count_8=c_8, count_9=c_9, count_10=c_10, score_truc=truc, score_cong=cong, score_tru=tru, note=note, total_score=total_val, start_date=start_date_val, end_date=end_date_val)
                        db_session.add(score); db_session.flush() 
                        log_details = f"Khởi tạo điểm mới lớp {branch.name} ({week_name}): {total_val}đ"
                        db_session.add(ActionLog(username=actor_username, full_name=actor_fullname, action_type="KHỞI TẠO ĐIỂM", details=log_details))
                        
                    db_session.query(WeeklyViolation).filter_by(weekly_score_id=score.id).delete()
                    for v in new_violations:
                        db_session.add(WeeklyViolation(weekly_score_id=score.id, violation_id=v['violation_id'], quantity=v['quantity'], student_name=v['student_name']))
                
                log_system_action("LƯU ĐIỂM TUẦN", f"Đã lưu và cập nhật bảng điểm {week_name}")
                flash(f"Đã lưu và cập nhật chính xác bảng điểm {week_name}!", "success")
                return redirect(url_for('weekly', week=week_name))

            week_param = request.args.get('week')
            if week_param:
                current_week = week_param
            else:
                all_weeks = db_session.query(WeeklyScore.week).join(Branch).filter(Branch.school_year_id == active_year.id if active_year else True).distinct().all()
                if all_weeks:
                    current_week = max([w[0] for w in all_weeks], key=lambda x: int(re.search(r'\d+', str(x)).group()) if re.search(r'\d+', str(x)) else 0)
                else:
                    latest_assign = db_session.query(Assignment).order_by(Assignment.week_number.desc()).first()
                    current_week = f"Tuần {latest_assign.week_number}" if latest_assign else "Tuần 1"

            branches_data = []; cat_list = []
            categories = db_session.query(ViolationCategory).filter_by(school_year_id=active_year.id).all() if active_year else []
            for c in categories: cat_list.append({"name": c.name, "points": float(c.penalty_points), "type": getattr(c, 'point_type', 'Điểm trừ')})
            categories_json = json.dumps(cat_list)

            start_date_str = ""; end_date_str = ""
            existing_score = db_session.query(WeeklyScore).join(Branch).filter(WeeklyScore.week == current_week, Branch.school_year_id == active_year.id if active_year else True).first()
            if existing_score and existing_score.start_date: start_date_str = existing_score.start_date; end_date_str = existing_score.end_date or ""
            else:
                try:
                    import datetime as dt
                    week_num = int(current_week.replace("Tuần ", "").strip())
                    assign_record = db_session.query(Assignment).filter(Assignment.week_number == week_num).first()
                    if assign_record and hasattr(assign_record, 'date') and assign_record.date:
                        py_date = assign_record.date; start_date_str = py_date.strftime("%Y-%m-%d")
                        day_of_week = py_date.weekday()
                        if day_of_week <= 5: days_to_add = 5 - day_of_week
                        else: days_to_add = 6
                        end_date = py_date + dt.timedelta(days=days_to_add)
                        end_date_str = end_date.strftime("%Y-%m-%d")
                except Exception as e: pass

            if active_year:
                branches = db_session.query(Branch).filter(Branch.school_year_id == active_year.id).all()
                branches.sort(key=lambda b: [int(t) if t.isdigit() else t.lower() for t in re.split(r'(\d+)', str(b.name))])
                for b in branches:
                    sc = db_session.query(WeeklyScore).filter_by(branch_id=b.id, week=current_week).first()
                    branches_data.append({'branch': b, 'score': sc})
                    
            settings = db_session.query(ScoreSettings).filter_by(school_year_id=active_year.id).first() if active_year else None
            return render_template('weekly.html', branches_data=branches_data, active_year=active_year, current_week=current_week, categories_json=categories_json, start_date=start_date_str, end_date=end_date_str, score_settings=settings)
    except Exception as e:
        flash(f"Lỗi nhập điểm tuần: {e}", "error")
        return redirect(url_for('dashboard'))
    
# ==========================================
# API: LƯU CẤU HÌNH BAREM ĐIỂM
# ==========================================
@app.route('/save_settings', methods=['POST'])
def save_settings():
    if session.get('role') not in ['Quản trị viên', 'Admin', 'Bí thư Đoàn trường', 'Bí thư']:
        flash("Chỉ Ban chấp hành Đoàn trường mới có quyền thay đổi Barem!", "error")
        return redirect(request.referrer or url_for('dashboard'))
        
    try:
        with session_scope() as db_session:
            active_year = db_session.query(SchoolYear).filter_by(is_active=True).first()
            if not active_year:
                flash("Chưa có năm học kích hoạt!", "error")
                return redirect(request.referrer or url_for('dashboard'))
                
            # Lấy dữ liệu từ form trên giao diện
            diem_8 = float(request.form.get('diem_8', 1.0))
            diem_9 = float(request.form.get('diem_9', 3.0))
            diem_10 = float(request.form.get('diem_10', 5.0))
            tuan_kha = float(request.form.get('tuan_kha', 20.0))
            tuan_tot = float(request.form.get('tuan_tot', 30.0))
            max_tot = int(request.form.get('max_tot', 14))
            max_mon = int(request.form.get('max_mon', 4))
            
            # Cập nhật vào DB
            settings = db_session.query(ScoreSettings).filter_by(school_year_id=active_year.id).first()
            if settings:
                settings.diem_8 = diem_8; settings.diem_9 = diem_9; settings.diem_10 = diem_10
                settings.diem_tuan_kha = tuan_kha; settings.diem_tuan_tot = tuan_tot
                settings.max_diem_tot = max_tot; settings.max_diem_mon = max_mon
            else:
                settings = ScoreSettings(
                    school_year_id=active_year.id,
                    diem_8=diem_8, diem_9=diem_9, diem_10=diem_10,
                    diem_tuan_kha=tuan_kha, diem_tuan_tot=tuan_tot,
                    max_diem_tot=max_tot, max_diem_mon=max_mon
                )
                db_session.add(settings)
                
            log_system_action("CẤU HÌNH", f"Cập nhật Barem điểm: Đ.10={diem_10}, Đ.9={diem_9}, Đ.8={diem_8}")
            flash("✅ Đã cập nhật cấu hình Barem điểm thành công! Hệ thống sẽ áp dụng Barem mới.", "success")
    except Exception as e:
        import traceback; traceback.print_exc()
        flash(f"Lỗi lưu Barem: {e}", "error")
        
    return redirect(request.referrer or url_for('weekly'))    
# ==========================================
# MODULE TẠO FILE EXCEL SỔ ĐEN (ĐÃ NÂNG CẤP TỔNG HỢP TOÀN TRƯỜNG & HIỂN THỊ NGÀY)
# ==========================================
@app.route('/preview_blacklist')
def preview_blacklist():
    try:
        with session_scope() as db_session:
            # Lấy bộ lọc thời gian (Mặc định là Cả năm học)
            time_filter = request.args.get('time_filter', 'Cả năm')
            active_year = db_session.query(SchoolYear).filter_by(is_active=True).first()
            
            if not active_year:
                flash("Chưa có năm học nào được kích hoạt!", "error")
                return redirect(url_for('dashboard'))
            
            # --- XÁC ĐỊNH DANH SÁCH TUẦN CẦN LỌC ---
            valid_weeks = []
            if time_filter == 'Cả năm':
                weeks_db = db_session.query(WeeklyScore.week).join(Branch).filter(Branch.school_year_id == active_year.id).distinct().all()
                valid_weeks = [w[0] for w in weeks_db]
            elif time_filter.startswith('Tháng') or time_filter.startswith('Học kỳ'):
                time_rec = db_session.query(MonthlyRecord).filter_by(school_year_id=active_year.id, month_name=time_filter).first()
                if time_rec and time_rec.weeks_used:
                    valid_weeks = [w.strip() for w in time_rec.weeks_used.split(',') if w.strip()]
            else:
                valid_weeks = [time_filter] # Lọc theo 1 Tuần đơn lẻ
                
            if not valid_weeks:
                valid_weeks = ['_NO_DATA_']

            # --- TRUY VẤN TOÀN BỘ LỖI TRONG CÁC TUẦN ĐÃ CHỌN ---
            raw_violations = db_session.query(WeeklyViolation, WeeklyScore, Branch, ViolationCategory)\
                .join(WeeklyScore, WeeklyViolation.weekly_score_id == WeeklyScore.id)\
                .join(Branch, WeeklyScore.branch_id == Branch.id)\
                .join(ViolationCategory, WeeklyViolation.violation_id == ViolationCategory.id)\
                .filter(
                    WeeklyScore.week.in_(valid_weeks), 
                    Branch.school_year_id == active_year.id
                ).all()
                
            # [THUẬT TOÁN TỔNG HỢP]: Dùng Dictionary để cộng dồn lỗi và gom Thẻ ngày
            import re
            summary_dict = {}
            
            for v, s, b, c in raw_violations:
                raw_val = str(v.student_name) if v.student_name else ""
                
                # Bóc tách thẻ ngày (VD: [T2], [T4]) đang cất giấu trong CSDL
                match_day = re.search(r'\[(T[2-7](?:\s*Chiều|\s*Chieu)?|CN)\]', raw_val, re.IGNORECASE)
                day_str = match_day.group(0).upper() if match_day else ""
                
                # Gỡ bỏ thẻ ngày để trả lại tên trắng sạch cho học sinh
                clean_names_str = raw_val.replace(day_str, "").strip() if day_str else raw_val
                
                if clean_names_str and clean_names_str != "":
                    raw_names = clean_names_str.replace(';', ',').split(',')
                    valid_names = [n.strip().title() for n in raw_names if n.strip()]
                    num_names = len(valid_names)
                    
                    # Chia đều lỗi nếu ghi chùm (VD: 3 em cùng vắng học -> Mỗi em 1 lỗi)
                    qty_per_student = max(1, v.quantity // num_names) if num_names > 0 else v.quantity
                    
                    for n_clean in valid_names:
                        key = (b.name, n_clean, c.name)
                        # Nếu là lỗi mới thì tạo giỏ chứa, có rồi thì cộng dồn
                        if key not in summary_dict:
                            summary_dict[key] = {'qty': 0, 'days': set()}
                            
                        summary_dict[key]['qty'] += qty_per_student
                        # Thêm thẻ ngày vào Set (Set sẽ tự động loại bỏ ngày trùng lặp)
                        if day_str: summary_dict[key]['days'].add(day_str)
                else:
                    # Bắt các Lỗi Tập Thể Lớp (Không có học sinh cụ thể)
                    key = (b.name, "Tập thể lớp", c.name)
                    if key not in summary_dict:
                        summary_dict[key] = {'qty': 0, 'days': set()}
                        
                    summary_dict[key]['qty'] += v.quantity
                    if day_str: summary_dict[key]['days'].add(day_str)
            
            # Đóng gói lại thành List để gửi ra giao diện
            violations = []
            for (b_name, stu_name, vio_name), data in summary_dict.items():
                # Ráp nối các thẻ ngày lại (VD: [T2] [T5])
                days_joined = " ".join(sorted(list(data['days'])))
                
                # Gắn chuỗi ngày vào đuôi tên Lỗi để hiển thị
                display_vio_name = f"{vio_name} {days_joined}" if days_joined else vio_name
                
                violations.append({
                    'branch_name': b_name,
                    'student_name': stu_name,
                    'violation_name': display_vio_name,
                    'quantity': data['qty']
                })
                    
            # Sắp xếp danh sách vi phạm: Ưu tiên Tên Lớp -> Số lần vi phạm nhiều nhất lên đầu
            violations.sort(key=lambda x: (x['branch_name'], -x['quantity']))
            
            # Tạo danh sách các mốc thời gian để làm Dropdown chọn bộ lọc trên giao diện
            time_options = ['Cả năm']
            sems = db_session.query(MonthlyRecord.month_name).filter(MonthlyRecord.school_year_id == active_year.id, MonthlyRecord.month_name.like('Học kỳ%')).distinct().all()
            months = db_session.query(MonthlyRecord.month_name).filter(MonthlyRecord.school_year_id == active_year.id, MonthlyRecord.month_name.like('Tháng%')).distinct().all()
            weeks = db_session.query(WeeklyScore.week).join(Branch).filter(Branch.school_year_id == active_year.id).distinct().all()
            
            time_options.extend(sorted([s[0] for s in sems]))
            time_options.extend(sorted([m[0] for m in months]))
            time_options.extend(sorted([w[0] for w in weeks], key=lambda x: int(re.search(r'\d+', x).group()) if re.search(r'\d+', x) else 0, reverse=True))

            return render_template('preview_blacklist.html', violations=violations, time_filter=time_filter, time_options=time_options)
    except Exception as e:
        import traceback; traceback.print_exc()
        flash(f"Lỗi xem trước sổ đen: {str(e)}", "error")
        return redirect(url_for('dashboard'))
    
# ==========================================
# MODULE XUẤT EXCEL SỔ ĐEN (ĐÃ NÂNG CẤP ĐÍNH KÈM THẺ NGÀY)
# ==========================================
@app.route('/export_blacklist', methods=['GET', 'POST'])
def export_blacklist():
    try:
        with session_scope() as db_session:
            # Nhận lệnh linh hoạt từ cả GET (nút nhấn) và POST (Form)
            time_filter = request.values.get('time_filter', 'Cả năm')
            active_year = db_session.query(SchoolYear).filter_by(is_active=True).first()
            
            if not active_year:
                flash("Chưa có năm học nào được kích hoạt!", "error")
                return redirect(url_for('dashboard'))
                
            valid_weeks = []
            if time_filter == 'Cả năm':
                weeks_db = db_session.query(WeeklyScore.week).join(Branch).filter(Branch.school_year_id == active_year.id).distinct().all()
                valid_weeks = [w[0] for w in weeks_db]
            elif time_filter.startswith('Tháng') or time_filter.startswith('Học kỳ'):
                time_rec = db_session.query(MonthlyRecord).filter_by(school_year_id=active_year.id, month_name=time_filter).first()
                if time_rec and time_rec.weeks_used:
                    valid_weeks = [w.strip() for w in time_rec.weeks_used.split(',') if w.strip()]
            else:
                valid_weeks = [time_filter]
                
            if not valid_weeks: valid_weeks = ['_NO_DATA_']

            raw_violations = db_session.query(WeeklyViolation, WeeklyScore, Branch, ViolationCategory)\
                .join(WeeklyScore, WeeklyViolation.weekly_score_id == WeeklyScore.id)\
                .join(Branch, WeeklyScore.branch_id == Branch.id)\
                .join(ViolationCategory, WeeklyViolation.violation_id == ViolationCategory.id)\
                .filter(WeeklyScore.week.in_(valid_weeks), Branch.school_year_id == active_year.id).all()
                
            import re
            summary_dict = {}
            for v, sc, b, c in raw_violations:
                raw_val = str(v.student_name) if v.student_name else ""
                
                # =========================================================
                # BÓC TÁCH: Lấy Thẻ ngày đang giấu trong tên HS ra ngoài
                # =========================================================
                match_day = re.search(r'\[(T[2-7](?:\s*Chiều|\s*Chieu)?|CN)\]', raw_val, re.IGNORECASE)
                day_str = match_day.group(0).upper() if match_day else ""
                
                clean_names_str = raw_val.replace(day_str, "").strip() if day_str else raw_val
                
                if clean_names_str and clean_names_str != "":
                    raw_names = clean_names_str.replace(';', ',').split(',')
                    valid_names = [n.strip().title() for n in raw_names if n.strip()]
                    num_names = len(valid_names)
                    qty_per_student = max(1, v.quantity // num_names) if num_names > 0 else v.quantity
                    
                    for n_clean in valid_names:
                        key = (b.name, n_clean, c.name)
                        if key not in summary_dict:
                            summary_dict[key] = {'qty': 0, 'days': set()}
                        summary_dict[key]['qty'] += qty_per_student
                        if day_str: summary_dict[key]['days'].add(day_str)
                else:
                    key = (b.name, "Tập thể lớp", c.name)
                    if key not in summary_dict:
                        summary_dict[key] = {'qty': 0, 'days': set()}
                    summary_dict[key]['qty'] += v.quantity
                    if day_str: summary_dict[key]['days'].add(day_str)
            
            violations = []
            for (b_name, stu_name, vio_name), data in summary_dict.items():
                # Ráp nối các thẻ ngày lại và đính vào đuôi tên lỗi
                days_joined = " ".join(sorted(list(data['days'])))
                display_vio_name = f"{vio_name} {days_joined}" if days_joined else vio_name
                
                violations.append({'branch_name': b_name, 'student_name': stu_name, 'violation_name': display_vio_name, 'quantity': data['qty']})
                
            violations.sort(key=lambda x: (x['branch_name'], -x['quantity']))
                
            if not violations:
                flash(f"Tuyệt vời! Không có dữ liệu vi phạm nào trong {time_filter}.", "success")
                return redirect(url_for('preview_blacklist', time_filter=time_filter))
                
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "So_Den_Toan_Truong"
            
            ws['A1'] = "ĐOÀN TRƯỜNG THPT THANH HÒA"
            ws['A1'].font = Font(name="Times New Roman", size=11, bold=True)
            ws['A3'] = f"TỔNG HỢP HỒ SƠ CÁ BIỆT (SỔ ĐEN) - {time_filter.upper()}"
            ws['A3'].font = Font(name="Times New Roman", size=14, bold=True)
            ws['A3'].alignment = Alignment(horizontal="center")
            ws.merge_cells('A3:E3')
            
            headers = ["STT", "Chi đoàn", "Họ và Tên Học Sinh", "Lỗi Vi Phạm", "Tổng Số Lần"]
            thin = Side(border_style="thin", color="000000")
            border = Border(left=thin, right=thin, top=thin, bottom=thin)
            
            for col, h in enumerate(headers, 1):
                c = ws.cell(row=5, column=col, value=h)
                c.font = Font(name="Times New Roman", size=11, bold=True)
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.border = border
                
            for idx, item in enumerate(violations, 1):
                row_idx = idx + 5
                c1 = ws.cell(row=row_idx, column=1, value=idx)
                c2 = ws.cell(row=row_idx, column=2, value=item['branch_name'])
                c3 = ws.cell(row=row_idx, column=3, value=item['student_name'])
                c4 = ws.cell(row=row_idx, column=4, value=item['violation_name'])
                c5 = ws.cell(row=row_idx, column=5, value=item['quantity'])
                
                for cell in [c1, c2, c3, c4, c5]:
                    cell.font = Font(name="Times New Roman", size=11)
                    cell.border = border
                c1.alignment = Alignment(horizontal="center")
                c2.alignment = Alignment(horizontal="center")
                c5.alignment = Alignment(horizontal="center")

                # Điểm nhấn: Bôi đỏ học sinh vi phạm từ 3 lần trở lên
                if item['quantity'] >= 3:
                    c5.font = Font(name="Times New Roman", size=11, bold=True, color="FF0000")
                    c3.font = Font(name="Times New Roman", size=11, bold=True, color="FF0000")
                
            ws.column_dimensions['A'].width = 6
            ws.column_dimensions['B'].width = 12
            ws.column_dimensions['C'].width = 25
            ws.column_dimensions['D'].width = 35
            ws.column_dimensions['E'].width = 15
            
            log_system_action("XUẤT EXCEL", f"Xuất Sổ đen Toàn trường - {time_filter}")
            out = io.BytesIO()
            wb.save(out)
            out.seek(0)
            
            return send_file(out, download_name=f"So_Den_{time_filter.replace(' ', '_')}.xlsx", as_attachment=True)
    except Exception as e:
        import traceback; traceback.print_exc() 
        flash(f"Lỗi xuất sổ đen: {str(e)}", "error")
        return redirect(url_for('preview_blacklist'))

# ==========================================
# MODULE: QUẢN LÝ NGÂN HÀNG LỖI
# ==========================================
@app.route('/violation-categories', methods=['GET'])
def violation_categories():
    try:
        with session_scope() as db_session:
            active_year = db_session.query(SchoolYear).filter_by(is_active=True).first()
            if not active_year:
                flash("Cần kích hoạt Năm học trước khi quản lý Thư viện lỗi!", "error")
                return redirect(url_for('dashboard'))
                
            categories = db_session.query(ViolationCategory).filter_by(school_year_id=active_year.id).order_by(ViolationCategory.name).all()
            return render_template('violation_categories.html', categories=categories, active_year=active_year)
    except Exception as e:
        flash(f"Lỗi truy cập thư viện lỗi: {e}", "error")
        return redirect(url_for('dashboard'))

@app.route('/add_violation', methods=['POST'])
def add_violation():
    try:
        with session_scope() as db_session:
            active_year = db_session.query(SchoolYear).filter_by(is_active=True).first()
            name = request.form.get('name', '').strip()
            try: points = float(request.form.get('penalty_points', 1.0))
            except ValueError: points = 1.0
            point_type = request.form.get('point_type', 'Điểm trừ')

            if name and active_year:
                exist = db_session.query(ViolationCategory).filter_by(name=name, school_year_id=active_year.id).first()
                if exist:
                    flash(f"Lỗi '{name}' đã tồn tại trong thư viện của năm học này!", "error")
                else:
                    new_cat = ViolationCategory(
                        name=name, penalty_points=points, 
                        point_type=point_type, school_year_id=active_year.id 
                    )
                    db_session.add(new_cat)
                    log_system_action("NGÂN HÀNG LỖI", f"Thêm quy định mới: {name} ({points} điểm)")
                    flash(f"Đã thêm lỗi '{name}' vào thư viện của {active_year.name}!", "success")
    except Exception as e:
        flash(f"Lỗi: {str(e)}", "error")
    return redirect(url_for('violation_categories'))

@app.route('/edit_violation/<int:id>', methods=['POST'])
def edit_violation(id):
    try:
        with session_scope() as db_session:
            cat = db_session.query(ViolationCategory).filter_by(id=id).first()
            if cat:
                cat.name = request.form.get('edit_name', cat.name).strip()
                try: cat.penalty_points = float(request.form.get('edit_penalty_points', cat.penalty_points))
                except: pass
                cat.point_type = request.form.get('edit_point_type', cat.point_type)
                log_system_action("NGÂN HÀNG LỖI", f"Cập nhật quy định lỗi: {cat.name}")
                flash(f"Đã cập nhật quy định cho lỗi '{cat.name}'!", "success")
    except Exception as e:
        flash(f"Lỗi cập nhật: {e}", "error")
    return redirect(url_for('violation_categories'))

@app.route('/delete_violation/<int:id>', methods=['POST'])
def delete_violation(id):
    try:
        with session_scope() as db_session:
            cat = db_session.query(ViolationCategory).filter_by(id=id).first()
            if cat:
                name = cat.name
                db_session.delete(cat)
                log_system_action("NGÂN HÀNG LỖI", f"Xóa quy định lỗi: {name}")
                flash(f"Đã xóa vĩnh viễn '{name}' khỏi hệ thống!", "success")
    except Exception as e:
        flash(f"Lỗi xóa lỗi: {e}", "error")
    return redirect(url_for('violation_categories'))

@app.route('/import_violations', methods=['POST'])
def import_violations():
    if 'excel_file' not in request.files:
        flash("Không tìm thấy file tải lên!", "error")
        return redirect(url_for('violation_categories'))
    
    file = request.files['excel_file']
    if file.filename == '':
        flash("Chưa chọn file nào!", "error")
        return redirect(url_for('violation_categories'))
        
    if file and (file.filename.endswith('.xlsx') or file.filename.endswith('.xls')):
        try:
            df = pd.read_excel(file)
            df.columns = df.columns.str.strip().str.upper() 
            
            with session_scope() as db_session:
                active_year = db_session.query(SchoolYear).filter_by(is_active=True).first()
                if not active_year:
                    flash("Chưa có năm học nào được kích hoạt!", "error")
                    return redirect(url_for('violation_categories'))
                    
                name_col = next((c for c in df.columns if "TÊN" in c or "LỖI" in c or "NỘI DUNG" in c), None)
                point_col = next((c for c in df.columns if "ĐIỂM" in c), None)
                type_col = next((c for c in df.columns if "LOẠI" in c or "CỘNG/TRỪ" in c), None)

                if not name_col or not point_col:
                    flash("Lỗi cấu trúc: File Excel cần có ít nhất cột 'Tên lỗi' và 'Số điểm'!", "error")
                    return redirect(url_for('violation_categories'))

                count_new = 0
                for index, row in df.iterrows():
                    name = str(row[name_col]).strip()
                    if not name or name.lower() == 'nan': continue

                    try: points = float(row[point_col])
                    except: points = 1.0

                    point_type = "Điểm trừ"
                    if type_col and pd.notna(row[type_col]):
                        val = str(row[type_col]).strip().lower()
                        if "cộng" in val or "+" in val:
                            point_type = "Điểm cộng"

                    exist = db_session.query(ViolationCategory).filter_by(name=name, school_year_id=active_year.id).first()
                    if not exist:
                        new_cat = ViolationCategory(
                            name=name, penalty_points=points, 
                            point_type=point_type, school_year_id=active_year.id
                        )
                        db_session.add(new_cat)
                        count_new += 1
                        
                log_system_action("NGÂN HÀNG LỖI", f"Đã nhập thành công {count_new} Quy định mới từ Excel")
                flash(f"Đã nhập thành công {count_new} Quy định mới từ file Excel!", "success")
        except Exception as e:
            flash(f"Lỗi đọc file Excel: {str(e)}", "error")
    else:
        flash("Vui lòng chọn file Excel hợp lệ (.xlsx, .xls)", "error")
        
    return redirect(url_for('violation_categories'))

# ==========================================
# MODULE: THI ĐUA THÁNG
# ==========================================
@app.route('/monthly', methods=['GET', 'POST'])
def monthly():
    try:
        with session_scope() as db_session:
            active_year = db_session.query(SchoolYear).filter(SchoolYear.is_active == True).first()
            available_weeks = []
            available_months = ["Tháng 9", "Tháng 10", "Tháng 11", "Tháng 12", "Tháng 1", "Tháng 2", "Tháng 3", "Tháng 4", "Tháng 5"]
            
            if active_year:
                weeks_db = db_session.query(WeeklyScore.week).join(Branch).filter(Branch.school_year_id == active_year.id).distinct().all()
                available_weeks = sorted([w[0] for w in weeks_db])

            monthly_data = {}
            selected_month = request.form.get('month') or request.args.get('month') or "Tháng 9"
            selected_weeks = request.form.getlist('weeks')

            used_by_other_months = set()
            current_month_weeks = set()
            special_periods = ["Thi đua chào mừng 20/11", "Thi đua chào mừng 26/3"]

            if active_year:
                all_records = db_session.query(MonthlyRecord).filter(
                    MonthlyRecord.school_year_id == active_year.id
                ).all()

                for rec in all_records:
                    if rec.weeks_used:
                        w_list = [w.strip() for w in rec.weeks_used.split(",") if w.strip()]
                        if rec.month_name == selected_month:
                            for w in w_list: current_month_weeks.add(w)
                        elif rec.month_name not in special_periods:
                            for w in w_list: used_by_other_months.add(w)

            if request.method == 'POST':
                if not selected_weeks:
                    flash("Vui lòng chọn ít nhất 1 tuần để tổng hợp điểm!", "error")
                elif active_year:
                    branches = db_session.query(Branch).filter(Branch.school_year_id == active_year.id).all()
                    
                    temp_groups = {}
                    for b in branches:
                        grp = b.group or "Nhóm 1"
                        if grp not in temp_groups:
                            temp_groups[grp] = []
                            
                        scores = db_session.query(WeeklyScore).filter(
                            WeeklyScore.branch_id == b.id,
                            WeeklyScore.week.in_(selected_weeks)
                        ).all()
                        
                        total_score = sum([(s.total_score or 0.0) for s in scores])
                        total_tru = sum([(s.score_tru or 0.0) for s in scores])   # Rút trích tổng lỗi
                        total_cong = sum([(s.score_cong or 0.0) for s in scores]) # Rút trích tổng thưởng
                        week_scores_dict = {s.week: (s.total_score or 0.0) for s in scores}
                        
                        temp_groups[grp].append({
                            'branch_id': b.id,
                            'branch_name': b.name,
                            'group': grp,
                            'gvcn': b.gvcn,
                            'total_score': total_score,
                            'total_tru': total_tru,
                            'total_cong': total_cong,
                            'weeks_count': len(scores),
                            'week_scores': week_scores_dict 
                        })
                    
                    for grp, lst in temp_groups.items():
                        # Áp dụng bộ lọc 4 lớp cho Thi đua Tháng
                        lst.sort(key=lambda x: (
                            -float(x['total_score']), 
                            float(x['total_tru']), 
                            -float(x['total_cong']),
                            x['branch_name']
                        ))
                        
                        current_rank = 1
                        for i, d in enumerate(lst):
                            if i > 0:
                                prev = lst[i-1]
                                if not (d['total_score'] == prev['total_score'] and 
                                        d['total_tru'] == prev['total_tru'] and 
                                        d['total_cong'] == prev['total_cong']):
                                    current_rank = i + 1
                            d['rank'] = current_rank
                        monthly_data[grp] = lst
                        
                    db_session.query(MonthlyRecord).filter(
                        MonthlyRecord.school_year_id == active_year.id, 
                        MonthlyRecord.month_name == selected_month
                    ).delete()
                    
                    w_str = ", ".join(selected_weeks)
                    for grp, lst in monthly_data.items():
                        for d in lst: 
                            db_session.add(MonthlyRecord(
                                school_year_id=active_year.id, 
                                month_name=selected_month, 
                                branch_id=d['branch_id'], 
                                total_score=d['total_score'], 
                                rank=d['rank'], 
                                weeks_used=w_str
                            ))
                    db_session.commit()
                    
                    log_system_action("LƯU ĐIỂM THÁNG", f"Đã tính toán và chốt sổ điểm {selected_month} (gộp từ: {w_str}).")
                    flash(f"✅ Đã lưu thành công dữ liệu xếp hạng {selected_month}!", "success")
            else:
                selected_weeks = list(current_month_weeks)

            user_role = session.get('role', '')
            is_admin = (user_role and ("Quản trị" in user_role or "Bí thư" in user_role or "Admin" in user_role))

            return render_template(
                'monthly.html', 
                available_weeks=available_weeks, 
                available_months=available_months,
                monthly_data=monthly_data,
                selected_month=selected_month, 
                selected_weeks=selected_weeks,
                used_by_other_months=used_by_other_months,
                is_admin=is_admin,
                active_year=active_year
            )
    except Exception as e:
        import traceback
        traceback.print_exc()
        flash(f"Lỗi phân hệ thi đua tháng: {e}", "error")
        return redirect(url_for('dashboard'))
    
# =========================================================================
# THUẬT TOÁN LÕI: TÍNH ĐIỂM TỐT CHUẨN XÁC TỪ BẢNG RAW_SCORES (BẢNG VÀNG)
# =========================================================================
def calculate_trimmed_details(raw_scores, branch_group, max_mon, max_tot):
    """Tính lại chi tiết số lượng điểm 10, 9, 8 dựa trên dữ liệu thô và barem động"""
    f10, f9, f8 = 0, 0, 0
    for r in raw_scores:
        try: c10, c9, c8 = int(r.c10 or 0), int(r.c9 or 0), int(r.c8 or 0)
        except ValueError: continue

        if "1" in str(branch_group): c8 = 0

        k10 = min(c10, max_mon)
        k9 = min(c9, max_mon - k10)
        k8 = min(c8, max_mon - k10 - k9)
        f10 += k10; f9 += k9; f8 += k8

    total = f10 + f9 + f8
    if total > max_tot:
        excess = total - max_tot
        cut_8 = min(f8, excess)
        f8 -= cut_8; excess -= cut_8
        if excess > 0:
            cut_9 = min(f9, excess)
            f9 -= cut_9; excess -= cut_9
        if excess > 0:
            f10 -= excess
    return f10, f9, f8

# ==========================================
# MODULE: BÁO CÁO & THỐNG KÊ
# ==========================================
@app.route('/reports', methods=['GET', 'POST'])
def reports():
    try:
        with session_scope() as db_session:
            years = db_session.query(SchoolYear).order_by(SchoolYear.id.desc()).all()
            active_year = next((y for y in years if y.is_active), None)

            selected_year_id = request.args.get('year_id', type=int)
            if not selected_year_id and active_year: 
                selected_year_id = active_year.id

            selected_month_filter = request.args.get('time_filter', 'Tất cả')
            active_tab = request.args.get('active_tab', 'tab-dashboard')

            chart_data = {
                "bar_labels": [], "bar_values": [], "bar_colors": [],
                "line_labels": [], "line_values": [],
                "pie_bad_labels": [], "pie_bad_values": [],
                "pie_good_labels": [], "pie_good_values": []
            }
            analysis_general = []
            analysis_details = []
            available_months = []

            if selected_year_id:
                months_db = db_session.query(MonthlyRecord.month_name).filter(
                    MonthlyRecord.school_year_id == selected_year_id,
                    MonthlyRecord.month_name.like('Tháng%')
                ).distinct().all()
                school_order = ["Tháng 9", "Tháng 10", "Tháng 11", "Tháng 12", "Tháng 1", "Tháng 2", "Tháng 3", "Tháng 4", "Tháng 5"]
                available_months = sorted([m[0] for m in months_db if m[0]], key=lambda x: school_order.index(x) if x in school_order else 99)

                # [BẢN VÁ LỖI TRÀN RAM]: Xác định rõ bảng gốc (select_from) để chặn đứng CROSS JOIN
                scores_raw = db_session.query(
                    Branch.name.label('Lớp'), Branch.group.label('Nhóm'),
                    WeeklyScore.week.label('Tuần'), WeeklyScore.total_score.label('Điểm'),
                    WeeklyScore.count_8.label('C8'), WeeklyScore.count_9.label('C9'), WeeklyScore.count_10.label('C10')
                ).select_from(WeeklyScore).join(Branch, WeeklyScore.branch_id == Branch.id)\
                 .filter(Branch.school_year_id == selected_year_id)
                
                vios_raw = db_session.query(
                    Branch.name.label('Lớp'), ViolationCategory.name.label('Lỗi'), 
                    WeeklyViolation.quantity.label('SL'), WeeklyScore.week.label('Tuần')
                ).select_from(WeeklyViolation)\
                 .join(ViolationCategory, ViolationCategory.id == WeeklyViolation.violation_id)\
                 .join(WeeklyScore, WeeklyScore.id == WeeklyViolation.weekly_score_id)\
                 .join(Branch, Branch.id == WeeklyScore.branch_id)\
                 .filter(Branch.school_year_id == selected_year_id)

                if selected_month_filter != 'Tất cả' and selected_month_filter in available_months:
                    month_rec = db_session.query(MonthlyRecord.weeks_used).filter(
                        MonthlyRecord.school_year_id == selected_year_id,
                        MonthlyRecord.month_name == selected_month_filter
                    ).first()
                    if month_rec and month_rec.weeks_used:
                        valid_weeks = [w.strip() for w in month_rec.weeks_used.split(',') if w.strip()]
                        scores_raw = scores_raw.filter(WeeklyScore.week.in_(valid_weeks))
                        vios_raw = vios_raw.filter(WeeklyScore.week.in_(valid_weeks))
                    else:
                        scores_raw = scores_raw.filter(WeeklyScore.week == 'NONE')

                df = pd.DataFrame(scores_raw.all())
                df_vios = pd.DataFrame(vios_raw.all())

                if not df.empty:
                    df['Điểm'] = pd.to_numeric(df['Điểm'], errors='coerce').fillna(0)
                    for col in ['C8', 'C9', 'C10']:
                        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
                        
                    def calc_tot_qty(row):
                        nhom = str(row['Nhóm'])
                        if "1" in nhom: return row['C9'] + row['C10']
                        elif "2" in nhom: return row['C8'] + row['C9'] + row['C10']
                        return row['C9'] + row['C10']
                    df['Điểm Tốt'] = df.apply(calc_tot_qty, axis=1)

                    sum_scores = df.groupby('Lớp')['Điểm'].sum().sort_values(ascending=False)
                    if len(sum_scores) >= 6:
                        top_bottom = pd.concat([sum_scores.head(3), sum_scores.tail(3)])
                        colors = ['#28a745']*3 + ['#dc3545']*3
                    else:
                        top_bottom = sum_scores
                        colors = ['#17a2b8'] * len(sum_scores)
                        
                    chart_data['bar_labels'] = top_bottom.index.tolist()
                    chart_data['bar_values'] = top_bottom.values.round(1).tolist()
                    chart_data['bar_colors'] = colors

                    trend_df = df.groupby('Tuần')['Điểm'].mean().reset_index()
                    trend_df['Tuần_Num'] = trend_df['Tuần'].str.extract(r'(\d+)').astype(float)
                    trend_df = trend_df.sort_values(by='Tuần_Num')
                    chart_data['line_labels'] = [str(w).replace("Tuần ", "T.") for w in trend_df['Tuần']]
                    chart_data['line_values'] = trend_df['Điểm'].round(1).tolist()

                    good_counts = df.groupby('Lớp')['Điểm Tốt'].sum().sort_values(ascending=False)
                    good_counts = good_counts[good_counts > 0]
                    if not good_counts.empty:
                        if len(good_counts) > 5:
                            top_good = good_counts.head(4)
                            top_good['Các lớp khác'] = good_counts.iloc[4:].sum()
                            good_counts = top_good
                        chart_data['pie_good_labels'] = good_counts.index.tolist()
                        chart_data['pie_good_values'] = good_counts.values.tolist()

                    df['Khối'] = df['Lớp'].str.extract(r'(\d+)')
                    khoi_sum_avg = df.groupby(['Khối', 'Lớp'])['Điểm'].sum().groupby('Khối').mean()
                    if not khoi_sum_avg.empty:
                        analysis_general.append(("Khối dẫn đầu phong trào", f"Khối {khoi_sum_avg.idxmax()} ({khoi_sum_avg.max():.1f}đ/lớp)", "Tuyên dương tinh thần khối"))
                    
                    if not sum_scores.empty:
                        analysis_general.append(("Chi đoàn Xuất sắc nhất", f"{sum_scores.idxmax()} ({int(sum_scores.max())}đ)", "Đề xuất biểu dương"))
                        
                    sum_by_group = df.groupby(['Nhóm', 'Lớp'])['Điểm'].sum().reset_index()
                    for gk, gn in [('1', 'Nhóm 1'), ('2', 'Nhóm 2')]:
                        g_df = sum_by_group[sum_by_group['Nhóm'].astype(str).str.contains(gk, na=False)]
                        if not g_df.empty:
                            best = g_df.loc[g_df['Điểm'].idxmax()]
                            analysis_general.append((f"Chi đoàn Tốt nhất ({gn})", f"{best['Lớp']} ({int(best['Điểm'])}đ)", "Khen thưởng theo nhóm"))
                            worst = g_df.loc[g_df['Điểm'].idxmin()]
                            analysis_general.append((f"Chi đoàn Yếu kém nhất ({gn})", f"{worst['Lớp']} ({int(worst['Điểm'])}đ)", "Nhắc nhở, đôn đốc chấn chỉnh"))

                    top_3_tot = good_counts.head(3)
                    if not top_3_tot.empty:
                        analysis_general.append(("Nhiều Điểm Tốt nhất", " | ".join([f"{cls} ({int(score)} lượt)" for cls, score in top_3_tot.items() if cls != 'Các lớp khác']), "Khích lệ phong trào"))
                        
                    overall_avg = df['Điểm'].mean()
                    trend = "Tốt" if overall_avg >= 85 else "Khá" if overall_avg >= 70 else "Thấp"
                    analysis_general.append(("Phong độ Toàn trường (ĐTB)", f"{overall_avg:.1f} điểm", f"Đánh giá: {trend}"))

                if not df_vios.empty:
                    df_vios['SL'] = pd.to_numeric(df_vios['SL'], errors='coerce').fillna(0)
                    bad_counts = df_vios.groupby('Lỗi')['SL'].sum().sort_values(ascending=False)
                    if not bad_counts.empty:
                        if len(bad_counts) > 5:
                            top_bad = bad_counts.head(4)
                            top_bad['Các lỗi khác'] = bad_counts.iloc[4:].sum()
                            bad_counts = top_bad
                        chart_data['pie_bad_labels'] = bad_counts.index.tolist()
                        chart_data['pie_bad_values'] = bad_counts.values.tolist()

                    class_error_summary = df_vios.groupby(['Lớp', 'Lỗi'])['SL'].sum().reset_index()
                    error_dict = {}; count_dict = {}
                    for _, row in class_error_summary.iterrows():
                        cls, err, qty = row['Lớp'], row['Lỗi'], int(row['SL'])
                        if cls not in error_dict: error_dict[cls] = []; count_dict[cls] = 0
                        error_dict[cls].append(f"{err} (x{qty})")
                        count_dict[cls] += qty

                    for cls, total_qty in sorted(count_dict.items(), key=lambda x: x[1], reverse=True):
                        analysis_details.append({"class": cls, "qty": total_qty, "details": " | ".join(error_dict[cls])})

            # --- 2. LOGIC BẢNG VÀNG HỌC TẬP ---
            honor_time = request.args.get('honor_time', 'Cả năm')
            honor_grade = request.args.get('honor_grade', 'Tất cả các khối')
            honor_data = []
            time_options_honor = ["Cả năm"]

            if selected_year_id:
                time_options_honor.extend(available_months)
                if len(time_options_honor) == 1: 
                    time_options_honor.extend(["Tháng 9", "Học kỳ 1", "Học kỳ 2"])

                max_tot, max_mon = 14, 4
                try:
                    settings = db_session.query(ScoreSettings).filter_by(school_year_id=selected_year_id).first()
                    if settings:
                        max_tot = int(getattr(settings, 'max_diem_tot', 14))
                        max_mon = int(getattr(settings, 'max_diem_mon', 4))
                except: pass

                valid_weeks_honor = []
                if honor_time == "Cả năm":
                    valid_weeks_honor = [w[0] for w in db_session.query(WeeklyScore.week).join(Branch).filter(Branch.school_year_id == selected_year_id).distinct().all()]
                elif "Học kì 1" in honor_time or "Học kỳ 1" in honor_time:
                    valid_weeks_honor = [f"Tuần {i}" for i in range(1, 19)]
                elif "Học kì 2" in honor_time or "Học kỳ 2" in honor_time:
                    valid_weeks_honor = [f"Tuần {i}" for i in range(19, 38)]
                else:
                    m_rec = db_session.query(MonthlyRecord).filter(MonthlyRecord.school_year_id == selected_year_id, MonthlyRecord.month_name == honor_time).first()
                    if m_rec and m_rec.weeks_used:
                        valid_weeks_honor = [w.strip() for w in m_rec.weeks_used.split(',') if w.strip()]
                
                if not valid_weeks_honor: 
                    valid_weeks_honor = ["_NO_DATA_"]

                branches_data = {}
                for b in db_session.query(Branch).filter(Branch.school_year_id == selected_year_id).all():
                    match = re.search(r'\d+', b.name)
                    branches_data[b.name] = {
                        "name": b.name, "khoi": f"Khối {match.group()}" if match else "Khác",
                        "c8": 0, "c9": 0, "c10": 0, "tuan_tot": 0, "total": 0, "group": str(b.group or "1")
                    }

                scores = db_session.query(WeeklyScore).join(Branch).filter(Branch.school_year_id == selected_year_id, WeeklyScore.week.in_(valid_weeks_honor)).all()
                for s in scores:
                    b_name = s.branch.name
                    b_group = branches_data[b_name]["group"]
                    
                    raw_scores = db_session.query(RawScore).filter_by(week=f"{s.week}_Y{selected_year_id}", branch_name=b_name).all()
                    if raw_scores:
                        f10, f9, f8 = calculate_trimmed_details(raw_scores, b_group, max_mon, max_tot)
                    else:
                        f10, f9, f8 = int(s.count_10 or 0), int(s.count_9 or 0), int(s.count_8 or 0)
                        if "1" in b_group: f8 = 0
                    
                    is_tot = 1 if s.week_rating and 'Tốt' in s.week_rating else 0
                    branches_data[b_name]["c10"] += f10
                    branches_data[b_name]["c9"] += f9
                    branches_data[b_name]["c8"] += f8
                    branches_data[b_name]["tuan_tot"] += is_tot
                    branches_data[b_name]["total"] += (f10 + f9 + f8)

                for d in branches_data.values():
                    if honor_grade == "Tất cả các khối" or d["khoi"] == honor_grade:
                        honor_data.append(d)
                honor_data.sort(key=lambda x: (x["total"], x["tuan_tot"], x["c10"], x["c9"], x["c8"]), reverse=True)

            return render_template(
                'reports.html', 
                years=years, 
                selected_year_id=selected_year_id,
                available_months=available_months,
                selected_month_filter=selected_month_filter,
                chart_data_json=json.dumps(chart_data),
                analysis_general=analysis_general,
                analysis_details=analysis_details,
                honor_data=honor_data,
                honor_time=honor_time,
                honor_grade=honor_grade,
                time_options_honor=time_options_honor,
                active_tab=active_tab
            )
    except Exception as e:
        import traceback; traceback.print_exc()
        flash(f"Lỗi tải báo cáo: {e}", "error")
        return redirect(url_for('dashboard'))

@app.route('/export_academic_honor', methods=['POST'])
def export_academic_honor():
    try:
        time_filter = request.form.get('time_filter', 'Cả năm')
        table_data_json = request.form.get('table_data')
        data = json.loads(table_data_json) if table_data_json else []

        if not data:
            flash("Không có dữ liệu để xuất!", "error")
            return redirect(url_for('reports', active_tab='tab-academic'))

        wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Bảng Vàng Học Tập"
        title_font = Font(name='Arial', size=14, bold=True, color="C00000")
        header_font = Font(name='Arial', size=11, bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="0054a6", end_color="0054a6", fill_type="solid")
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        center_align = Alignment(horizontal="center", vertical="center")
        
        ws.merge_cells('A1:H1')
        ws['A1'] = f"BẢNG VÀNG THÀNH TÍCH HỌC TẬP - {time_filter.upper()}"
        ws['A1'].font = title_font; ws['A1'].alignment = center_align
        
        headers = ["STT", "Tên Chi đoàn", "Khối", "Tổng Điểm 8", "Tổng Điểm 9", "Tổng Điểm 10", "Tổng Tuần Loại Tốt", "TỔNG ĐIỂM TỐT"]
        for c, h in enumerate(headers, 1):
            cell = ws.cell(row=3, column=c, value=h)
            cell.font = header_font; cell.fill = header_fill; cell.alignment = center_align; cell.border = thin_border
            
        for r, d in enumerate(data):
            row_vals = [r+1, d['name'], d['khoi'], d['c8'], d['c9'], d['c10'], d['tuan_tot'], d['total']]
            for c, val in enumerate(row_vals, 1):
                cell = ws.cell(row=4+r, column=c, value=val)
                cell.alignment = center_align; cell.border = thin_border
                if c in [7, 8]: cell.font = Font(color="C00000", bold=True)
                
        for col, w in zip(['A','B','C','D','E','F','G','H'], [6, 15, 12, 15, 15, 15, 20, 18]):
            ws.column_dimensions[col].width = w
            
        log_system_action("XUẤT EXCEL", f"Xuất Bảng Vàng Học Tập - {time_filter}")
        out = io.BytesIO(); wb.save(out); out.seek(0)
        return send_file(out, download_name=f"Bang_Vang_{time_filter.replace(' ', '_')}.xlsx", as_attachment=True)
    except Exception as e:
        flash(f"Lỗi xuất Excel: {e}", "error")
        return redirect(url_for('reports', active_tab='tab-academic'))

@app.route('/export_chronic_violations', methods=['POST'])
def export_chronic_violations():
    try:
        with session_scope() as db_session:
            year_id = request.form.get('export_year_id', type=int)
            month_name = request.form.get('export_month')
            active_year = db_session.query(SchoolYear).filter_by(id=year_id).first()
            if not active_year or not month_name:
                flash("Thiếu thông tin xuất báo cáo!", "error"); return redirect(url_for('reports'))

            month_rec = db_session.query(MonthlyRecord).filter(MonthlyRecord.school_year_id == year_id, MonthlyRecord.month_name == month_name).first()
            if not month_rec or not month_rec.weeks_used:
                flash(f"Chưa có dữ liệu chốt tháng cho {month_name}!", "error"); return redirect(url_for('reports'))
                
            valid_weeks = [w.strip() for w in month_rec.weeks_used.split(',') if w.strip()]
            branches = db_session.query(Branch).filter(Branch.school_year_id == year_id).all()
            b_ids = [b.id for b in branches]
            
            scores = db_session.query(WeeklyScore).filter(WeeklyScore.branch_id.in_(b_ids), WeeklyScore.week.in_(valid_weeks)).all()
            
            score_map = {}
            for s in scores:
                diem_tru = float(s.score_tru or 0.0)
                if diem_tru > 0:
                    if s.branch_id not in score_map: score_map[s.branch_id] = {'tru': 0, 'weeks': 0}
                    score_map[s.branch_id]['tru'] += diem_tru
                    score_map[s.branch_id]['weeks'] += 1

            export_data = []
            for b in branches:
                if b.id in score_map:
                    d = score_map[b.id]
                    if d['weeks'] >= 2 or d['tru'] >= 5:
                        export_data.append([0, b.name, d['weeks'], d['tru'], "Vi phạm nhiều lần/Nghiêm trọng"])

            if not export_data:
                flash(f"Tuyệt vời! Không có lớp nào bị cảnh báo vi phạm trong {month_name}!", "success"); return redirect(url_for('reports'))

            export_data.sort(key=lambda x: x[3], reverse=True)

            wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Cảnh báo Vi phạm"
            ws.merge_cells('A1:E1'); ws['A1'] = f"BÁO CÁO CẢNH BÁO VI PHẠM NỀ NẾP - {month_name.upper()}"; ws['A1'].font = Font(size=14, bold=True); ws['A1'].alignment = Alignment(horizontal="center")
            ws.merge_cells('A2:E2'); ws['A2'] = f"(Gộp dữ liệu từ: {month_rec.weeks_used})"; ws['A2'].font = Font(italic=True); ws['A2'].alignment = Alignment(horizontal="center")
            
            bd = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            headers = ["STT", "Lớp", "Số tuần vi phạm", "Tổng điểm bị trừ", "Đánh giá sơ bộ"]
            for col, h in enumerate(headers, 1):
                c = ws.cell(row=4, column=col, value=h); c.font = Font(bold=True); c.alignment = Alignment(horizontal="center"); c.border = bd
                
            for idx, row in enumerate(export_data, 1):
                row[0] = idx
                for col, val in enumerate(row, 1):
                    c = ws.cell(row=idx+4, column=col, value=val); c.border = bd
                    if col in [1, 2, 3, 4]: c.alignment = Alignment(horizontal="center")

            ws.column_dimensions['B'].width = 15; ws.column_dimensions['C'].width = 18; ws.column_dimensions['D'].width = 18; ws.column_dimensions['E'].width = 35
            
            log_system_action("XUẤT EXCEL", f"Xuất báo cáo Cảnh báo Vi phạm - {month_name}")
            out = io.BytesIO(); wb.save(out); out.seek(0)
            return send_file(out, download_name=f"Canh_Bao_Vi_Pham_{month_name.replace(' ', '_')}.xlsx", as_attachment=True)
    except Exception as e:
        flash(f"Lỗi: {e}", "error"); return redirect(url_for('reports'))

@app.route('/export_monthly_summary', methods=['POST'])
def export_monthly_summary():
    try:
        with session_scope() as db_session:
            year_id = request.form.get('export_year_id', type=int)
            month_name = request.form.get('export_month')
            active_year = db_session.query(SchoolYear).filter_by(id=year_id).first()
            if not active_year or not month_name:
                flash("Thiếu thông tin xuất báo cáo!", "error"); return redirect(url_for('reports'))

            month_rec = db_session.query(MonthlyRecord).filter(MonthlyRecord.school_year_id == year_id, MonthlyRecord.month_name == month_name).first()
            if not month_rec or not month_rec.weeks_used:
                flash(f"Chưa có dữ liệu chốt tháng cho {month_name}!", "error"); return redirect(url_for('reports'))
                
            valid_weeks = [w.strip() for w in month_rec.weeks_used.split(',') if w.strip()]
            branches = db_session.query(Branch).filter(Branch.school_year_id == year_id).all()
            b_ids = [b.id for b in branches]
            
            scores = db_session.query(WeeklyScore).filter(WeeklyScore.branch_id.in_(b_ids), WeeklyScore.week.in_(valid_weeks)).all()
            
            score_map = {}
            for s in scores:
                if s.branch_id not in score_map: score_map[s.branch_id] = {'tot': 0, 'tru': 0, 'tong': 0, 'count': 0}
                sl_tot = int(s.count_8 or 0) + int(s.count_9 or 0) + int(s.count_10 or 0)
                score_map[s.branch_id]['tot'] += sl_tot
                score_map[s.branch_id]['tru'] += float(s.score_tru or 0.0)
                score_map[s.branch_id]['tong'] += float(s.total_score or 0.0)
                score_map[s.branch_id]['count'] += 1

            export_data = []
            for b in branches:
                if b.id in score_map:
                    d = score_map[b.id]
                    avg = d['tong'] / d['count'] if d['count'] > 0 else 0
                    export_data.append([0, b.name, b.group, d['tot'], d['tru'], round(avg, 1)])

            if not export_data:
                flash(f"Chưa có điểm số nào!", "error"); return redirect(url_for('reports'))

            export_data.sort(key=lambda x: x[5], reverse=True)

            wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Tổng hợp Thi đua"
            ws.merge_cells('A1:F1'); ws['A1'] = f"TỔNG HỢP KẾT QUẢ THI ĐUA - {month_name.upper()}"; ws['A1'].font = Font(size=14, bold=True); ws['A1'].alignment = Alignment(horizontal="center")
            ws.merge_cells('A2:F2'); ws['A2'] = f"(Gộp dữ liệu từ: {month_rec.weeks_used})"; ws['A2'].font = Font(italic=True); ws['A2'].alignment = Alignment(horizontal="center")
            
            bd = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            headers = ["STT", "Lớp", "Nhóm", "Tổng Điểm Tốt", "Tổng Điểm Bị Trừ", "Điểm TB Tháng"]
            for col, h in enumerate(headers, 1):
                c = ws.cell(row=4, column=col, value=h); c.font = Font(bold=True); c.alignment = Alignment(horizontal="center"); c.border = bd
                
            for idx, row in enumerate(export_data, 1):
                row[0] = idx
                for col, val in enumerate(row, 1):
                    c = ws.cell(row=idx+4, column=col, value=val); c.border = bd
                    c.alignment = Alignment(horizontal="center")

            ws.column_dimensions['B'].width = 12; ws.column_dimensions['C'].width = 12; ws.column_dimensions['D'].width = 18; ws.column_dimensions['E'].width = 18; ws.column_dimensions['F'].width = 18
            
            log_system_action("XUẤT EXCEL", f"Xuất Tổng hợp Kế quả Thi đua - {month_name}")
            out = io.BytesIO(); wb.save(out); out.seek(0)
            return send_file(out, download_name=f"Tong_Hop_Thi_Dua_{month_name.replace(' ', '_')}.xlsx", as_attachment=True)
    except Exception as e:
        flash(f"Lỗi: {e}", "error"); return redirect(url_for('reports'))

# ==========================================
# MODULE: TRA CỨU CHI ĐOÀN (CLASS DASHBOARD)
# ==========================================
@app.route('/class-dashboard', methods=['GET', 'POST'])
def class_dashboard():
    try:
        with session_scope() as db_session:
            active_year = db_session.query(SchoolYear).filter(SchoolYear.is_active == True).first()
            
            # [NÂNG CẤP LÕI]: Khoanh vùng dữ liệu theo Quyền đăng nhập
            user_role = session.get('role', '')
            session_username = session.get('username', '').strip().upper()
            is_gvcn = (user_role == 'Giáo viên chủ nhiệm')
            
            branches = []
            if active_year:
                if is_gvcn:
                    branches = db_session.query(Branch).filter(Branch.name == session_username, Branch.school_year_id == active_year.id).all()
                else:
                    branches = db_session.query(Branch).filter(Branch.school_year_id == active_year.id).all()
                    
                    # [THUẬT TOÁN SẮP XẾP TỰ NHIÊN - NATURAL SORT]
                    import re
                    branches.sort(key=lambda b: [int(t) if t.isdigit() else t.lower() for t in re.split(r'(\d+)', str(b.name))])
                    
            selected_branch_id = request.args.get('branch_id', type=int)
            if not selected_branch_id and branches:
                selected_branch_id = branches[0].id
                
            selected_branch = None
            weekly_scores = []
            assignments = []
            monitoring_assignments = []
            warning_students = []
            appeal_records = []
            
            # =========================================================================
            # [NÂNG CẤP]: LẤY DỮ LIỆU NGÂN HÀNG LỖI ĐỂ GVCN TRA CỨU
            # =========================================================================
            violation_bank = []
            if active_year:
                violation_bank = db_session.query(ViolationCategory).filter_by(
                    school_year_id=active_year.id
                ).order_by(ViolationCategory.point_type.desc(), ViolationCategory.name).all()
            # =========================================================================
            
            if selected_branch_id:
                if is_gvcn and branches and selected_branch_id != branches[0].id:
                    selected_branch_id = branches[0].id
                    
                selected_branch = db_session.query(Branch).filter_by(id=selected_branch_id).first()
                if selected_branch:
                    group_val = selected_branch.group or "Nhóm 1"
                    
                    # [VÁ LỖI]: Lấy dữ liệu và sắp xếp theo số Tuần giảm dần (Tuần mới nhất lên đầu)
                    weekly_scores_db = db_session.query(WeeklyScore).filter_by(branch_id=selected_branch.id).all()
                    import re
                    weekly_scores_db.sort(key=lambda x: int(re.search(r'\d+', str(x.week)).group()) if x.week and re.search(r'\d+', str(x.week)) else 0, reverse=True)
                    
                    # =========================================================================
                    # [BẢN VÁ THUẬT TOÁN BÁO ĐỘNG]: CỘNG DỒN TỪ ĐẦU NĂM & BỎ QUA "ẢNH MINH CHỨNG"
                    # =========================================================================
                    if weekly_scores_db:
                        student_viol_counts = {} # Đưa bộ đếm ra ngoài vòng lặp tuần để cộng dồn
                        for score in weekly_scores_db:
                            for viol in score.violations:
                                # [KHIÊN BẢO VỆ]: Bỏ qua lỗi có tên "Ảnh minh chứng"
                                cat = db_session.query(ViolationCategory).filter_by(id=viol.violation_id).first()
                                if cat and "ảnh minh chứng" in cat.name.lower():
                                    continue

                                if viol.student_name and str(viol.student_name).strip() != "":
                                    raw_names = str(viol.student_name).replace(';', ',').split(',')
                                    names = [n.strip().upper() for n in raw_names if n.strip()]
                                    
                                    # [THUẬT TOÁN CHIA ĐỀU BỘ ĐẾM]
                                    num_names = len(names)
                                    total_qty = int(viol.quantity) if viol.quantity else 1
                                    qty_per_student = max(1, total_qty // num_names) if num_names > 0 else total_qty
                                    
                                    for name in names:
                                        student_viol_counts[name] = student_viol_counts.get(name, 0) + qty_per_student
                        
                        # Phân loại mức độ vi phạm cho lớp
                        for name, count in student_viol_counts.items():
                            if count >= 3:
                                if count >= 5:
                                    badge_class = "danger"
                                    badge_label = "Báo Động Đỏ"
                                else:
                                    badge_class = "warning text-dark"
                                    badge_label = "Cảnh Báo Vàng"
                                    
                                warning_students.append({
                                    'name': name.title(),
                                    'count': count,
                                    'badge_class': badge_class,
                                    'badge_label': badge_label
                                })
                        
                        warning_students.sort(key=lambda x: x['count'], reverse=True)
                    # =========================================================================
                    
                    # [BẢN VÁ MÚI GIỜ]: Ép giờ VN để khóa số lượt gửi phúc khảo trong ngày
                    from datetime import datetime, timezone, timedelta
                    vn_tz = timezone(timedelta(hours=7))
                    today_str = datetime.now(vn_tz).strftime("%d/%m/%Y")
                    
                    for sc in weekly_scores_db:
                        all_in_week = db_session.query(WeeklyScore).join(Branch).filter(
                            WeeklyScore.week == sc.week,
                            Branch.school_year_id == active_year.id
                        ).all()
                        
                        same_group_scores = [s for s in all_in_week if (s.branch.group or "Nhóm 1") == group_val]
                        
                        # =======================================================
                        # [THUẬT TOÁN ĐỒNG HẠNG TIE-BREAKER CHUẨN XÁC]
                        # =======================================================
                        same_group_scores.sort(key=lambda x: (
                            -float(x.total_score or 0),   # Ưu tiên 1: Tổng điểm (Từ cao xuống thấp)
                            float(x.score_tru or 0),      # Ưu tiên 2: Ít điểm trừ vi phạm hơn sẽ xếp trên
                            -float(x.score_cong or 0),    # Ưu tiên 3: Nhiều điểm thưởng hơn sẽ xếp trên
                            x.branch.name                 # Ưu tiên 4: Cùng điểm thì xếp hạng theo Tên Lớp (A-Z)
                        ))
                        
                        rk = 1
                        for i, s in enumerate(same_group_scores):
                            if i > 0:
                                prev = same_group_scores[i-1]
                                # Phải hoàn toàn giống nhau 3 hệ số mới được cấp Đồng Hạng
                                if not (float(s.total_score or 0) == float(prev.total_score or 0) and 
                                        float(s.score_tru or 0) == float(prev.score_tru or 0) and 
                                        float(s.score_cong or 0) == float(prev.score_cong or 0)):
                                    rk = i + 1 # Nhảy bậc xếp hạng (VD: 1, 2, 2, 4)
                            if s.branch_id == selected_branch.id: break
                        
                        so_luong_diem_tot = int(sc.count_9 or 0) + int(sc.count_10 or 0)
                        if "2" in str(group_val): so_luong_diem_tot += int(sc.count_8 or 0)
                            
                        # =========================================================================
                        # [NÂNG CẤP LÕI]: MỞ KHÓA NÚT GIAO DIỆN VÀ BÓC TÁCH HỒ SƠ PHÚC KHẢO
                        # =========================================================================
                        has_appealed_today = False
                        if sc.appeal_reason:
                            count_today = sc.appeal_reason.count(f"[{today_str}")
                            if count_today >= 2:
                                has_appealed_today = True
                                
                            # --- [BỔ SUNG]: BÓC TÁCH DỮ LIỆU CHO TAB HỒ SƠ PHÚC KHẢO ---
                            
                            pattern = r'\[\d{2}/\d{2}/\d{4} \d{2}:\d{2}\]'
                            timestamps = re.findall(pattern, sc.appeal_reason)
                            segments = re.split(pattern, sc.appeal_reason)[1:] 
                            
                            # Duyệt ngược để khiếu nại mới nhất lên đầu danh sách
                            for i in range(len(timestamps)-1, -1, -1):
                                time_str = timestamps[i].strip('[]')
                                content = segments[i].strip().strip('|').strip()
                                
                                errors_part = ""
                                reason_part = content
                                
                                # Cắt chuỗi để lấy riêng phần Lỗi và phần Lý do/Minh chứng
                                match = re.search(r'Phúc khảo các lỗi:\s*(.*?)\s*\|\s*Lý do:(.*)', content, re.IGNORECASE)
                                if match:
                                    errors_part = match.group(1).strip()
                                    reason_part = match.group(2).strip()
                                    
                                    # [TÍNH NĂNG MỚI]: LÀM ĐẸP CHUỖI HIỂN THỊ DẠNG BULLET CHO BCH
                                    errors_html = errors_part.replace('] & [', '<br>• ').replace('[', '').replace(']', '')
                                    if errors_html and not errors_html.startswith('• '):
                                        errors_html = '• ' + errors_html
                                    errors_part = errors_html
                                
                                status_text = "Đang chờ xử lý"
                                badge_class = "warning text-dark"
                                if sc.appeal_response:
                                    if "ĐÃ DUYỆT" in sc.appeal_response:
                                        status_text = "Đã duyệt"
                                        badge_class = "success"
                                    elif "TỪ CHỐI" in sc.appeal_response:
                                        status_text = "Từ chối"
                                        badge_class = "danger"
                                
                                appeal_records.append({
                                    'week': sc.week,
                                    'time': time_str,
                                    'errors_raw': errors_part,
                                    'reason': reason_part,
                                    'status': status_text,
                                    'badge': badge_class,
                                    'response': sc.appeal_response if sc.appeal_response else ""
                                })
                            # -----------------------------------------------------------
                        # =========================================================================
                            
                        weekly_scores.append({
                            'score_id': sc.id,             
                            'is_locked': getattr(sc, 'is_locked', False),     
                            'is_appealed': getattr(sc, 'is_appealed', False), 
                            'has_appealed_today': has_appealed_today,
                            'appeal_reason': getattr(sc, 'appeal_reason', ''),
                            'appeal_response': getattr(sc, 'appeal_response', ''),
                            'week': sc.week,
                            'rating': sc.week_rating or "-",
                            'score_truc': sc.score_truc or 100,
                            'diem_tot': so_luong_diem_tot,
                            'score_cong': sc.score_cong or 0,
                            'score_tru': sc.score_tru or 0,
                            'total_score': sc.total_score or 0,
                            'note': sc.note or "",
                            'rank': rk,
                            'is_appeal_expired': False, # [BẢN VÁ]: Tắt vĩnh viễn khóa 3 ngày của CSDL
                            'evidence_image': getattr(sc, 'evidence_image', None)
                        })
                    
                    red_star_ids = [rs.id for rs in selected_branch.red_stars] if hasattr(selected_branch, 'red_stars') else []
                    if not red_star_ids: 
                        red_star_ids = [rs.id for rs in db_session.query(RedStar).filter_by(branch_id=selected_branch.id).all()]
                    
                    import json, os
                    base_dir = os.path.dirname(os.path.abspath(__file__))
                    zones_map = {}
                    config_path = os.path.join(base_dir, "config", "class_zones.json")
                    if os.path.exists(config_path):
                        with open(config_path, "r", encoding="utf-8") as f:
                            try: zones_map = json.load(f)
                            except: pass
                    # =========================================================================
                    # [BẢN VÁ TỐI ƯU]: XỬ LÝ LỊCH PHÂN CÔNG ĐI TRỰC CHO TAB 3
                    # =========================================================================
                    if red_star_ids:
                        raw_assignments = db_session.query(Assignment).filter(Assignment.red_star_id.in_(red_star_ids)).order_by(Assignment.week_number.desc(), Assignment.shift).all()
                        
                        import os, json
                        from datetime import datetime, timedelta
                        base_dir = os.path.dirname(os.path.abspath(__file__))
                        zones_map = {}
                        config_path = os.path.join(base_dir, "config", "class_zones.json")
                        if os.path.exists(config_path):
                            with open(config_path, "r", encoding="utf-8") as f:
                                try: zones_map = json.load(f)
                                except: pass

                        for assign in raw_assignments:
                            # 1. XỬ LÝ LỖI NGÀY THÁNG SQLITE (Bắt chuẩn chuỗi String và Date)
                            date_range_str = ""
                            raw_date = getattr(assign, 'date', None)
                            if raw_date:
                                try:
                                    if isinstance(raw_date, str):
                                        date_str_clean = raw_date.split()[0]
                                        py_date = datetime.strptime(date_str_clean, '%Y-%m-%d').date()
                                    else:
                                        py_date = raw_date
                                        
                                    start_d = py_date.strftime("%d/%m/%Y")
                                    day_of_week = py_date.weekday()
                                    days_to_add = 5 - day_of_week if day_of_week <= 5 else 6
                                    end_date = py_date + timedelta(days=days_to_add)
                                    end_d = end_date.strftime("%d/%m/%Y")
                                    date_range_str = f"({start_d} đến {end_d})"
                                except Exception as e:
                                    pass

                            # 2. XỬ LÝ LỖI HOA/THƯỜNG KHI DỊCH TÊN LỚP (Quét không phân biệt hoa thường)
                            area_name = assign.duty_area.name if assign.duty_area else ""
                            classes_in_zone = []
                            for k, v in zones_map.items():
                                if str(k).strip().lower() == str(area_name).strip().lower():
                                    classes_in_zone = v
                                    break
                            task_display = ", ".join(classes_in_zone) if classes_in_zone else area_name

                            # 3. ĐÓNG GÓI CHUẨN XÁC TÊN BIẾN CHO GIAO DIỆN HTML
                            assignments.append({
                                'week_number': assign.week_number,
                                'red_star_name': assign.red_star.full_name if assign.red_star else "Khuyết",
                                'shift': assign.shift,
                                'date_range_str': date_range_str,
                                'task_display': task_display
                            })
                    # =========================================================================
                    # =========================================================================
                    
                    target_zone_names = [zone for zone, classes in zones_map.items() if isinstance(classes, list) and str(selected_branch.name or '').strip().upper() in [str(c).strip().upper() for c in classes]]
                    if target_zone_names: 
                        raw_monitoring_assignments = db_session.query(Assignment).join(DutyArea).join(RedStar).join(Branch).filter(DutyArea.name.in_(target_zone_names), Branch.school_year_id == active_year.id).order_by(Assignment.week_number.desc(), Assignment.shift).all()
                        
                        score_lookup = {sc.week: float(sc.total_score or 100.0) for sc in weekly_scores_db}
                        for mon_asm in raw_monitoring_assignments:
                            wk = f"Tuần {mon_asm.week_number}"
                            monitoring_assignments.append({
                                'assignment': mon_asm,
                                'week_score': score_lookup.get(wk, 100.0)
                            })

            return render_template('class_dashboard.html', 
                                   active_year=active_year, 
                                   branches=branches, 
                                   selected_branch=selected_branch, 
                                   weekly_scores=weekly_scores, 
                                   assignments=assignments, 
                                   monitoring_assignments=monitoring_assignments,
                                   warning_students=warning_students,
                                   violation_bank=violation_bank,
                                   appeal_records=appeal_records,
                                   is_gvcn=is_gvcn 
            )
    except Exception as e:
        import traceback; traceback.print_exc()
        flash(f"Lỗi tra cứu chi đoàn: {e}", "error")
        return redirect(url_for('dashboard'))
    
# ==========================================
# MODULE: TRANG XEM TRƯỚC HỒ SƠ LỚP MỚI
# ==========================================
@app.route('/preview_class_dashboard/<int:branch_id>')
def preview_class_dashboard(branch_id):
    try:
        with session_scope() as db_session:
            active_year = db_session.query(SchoolYear).filter_by(is_active=True).first()
            if not active_year: return redirect(url_for('class_dashboard'))

            branch = db_session.query(Branch).filter_by(id=branch_id).first()
            if not branch: return redirect(url_for('class_dashboard'))

            group_val = branch.group or "Nhóm 1"
            weekly_scores = []
            weekly_scores_db = db_session.query(WeeklyScore).filter_by(branch_id=branch.id).order_by(WeeklyScore.id).all()
            for sc in weekly_scores_db:
                all_in_week = db_session.query(WeeklyScore).join(Branch).filter(
                    WeeklyScore.week == sc.week, Branch.school_year_id == active_year.id
                ).all()
                same_group_scores = [s for s in all_in_week if (s.branch.group or "Nhóm 1") == group_val]
                same_group_scores.sort(key=lambda x: float(x.total_score or 0), reverse=True)
                rk = 1
                for i, s in enumerate(same_group_scores):
                    if i > 0 and float(s.total_score or 0) < float(same_group_scores[i-1].total_score or 0): rk = i + 1
                    if s.branch_id == branch.id: break
                
                # [ĐÃ NÂNG CẤP LẠI]: Tính Tổng số con điểm Tốt (Không nhân hệ số)
                so_luong_diem_tot = int(sc.count_9 or 0) + int(sc.count_10 or 0)
                if "2" in str(group_val): so_luong_diem_tot += int(sc.count_8 or 0)
                
                weekly_scores.append({
                    'week': sc.week, 'rating': sc.week_rating or "-", 'score_truc': sc.score_truc or 100,
                    'diem_tot': so_luong_diem_tot, 'score_cong': sc.score_cong or 0, 'score_tru': sc.score_tru or 0,
                    'total_score': sc.total_score or 0, 'note': sc.note or "", 'rank': rk
                })
            
            red_star_ids = [rs.id for rs in db_session.query(RedStar).filter_by(branch_id=branch.id).all()]
            assignments = []
            if red_star_ids:
                assignments = db_session.query(Assignment).filter(Assignment.red_star_id.in_(red_star_ids)).order_by(Assignment.week_number.desc(), Assignment.shift).all()

            import json, os
            base_dir = os.path.dirname(os.path.abspath(__file__))
            zones_map = {}
            monitoring_assignments = []
            if os.path.exists(os.path.join(base_dir, "config", "class_zones.json")):
                with open(os.path.join(base_dir, "config", "class_zones.json"), "r", encoding="utf-8") as f:
                    try: zones_map = json.load(f)
                    except: pass
                    
            target_zone_names = []
            current_branch_clean = str(branch.name or '').strip().upper()
            for zone, classes in zones_map.items():
                if isinstance(classes, list):
                    if current_branch_clean in [str(c).strip().upper() for c in classes]:
                        target_zone_names.append(zone)
                        
            if target_zone_names:
                monitoring_assignments = db_session.query(Assignment).join(DutyArea).join(RedStar).join(Branch).filter(
                    DutyArea.name.in_(target_zone_names), Branch.school_year_id == active_year.id
                ).order_by(Assignment.week_number.desc(), Assignment.shift).all()
                
            return render_template('preview_class_dashboard.html', branch=branch, weekly_scores=weekly_scores, assignments=assignments, monitoring_assignments=monitoring_assignments)
    except Exception as e:
        flash(f"Lỗi xem trước: {e}", "error")
        return redirect(url_for('class_dashboard'))

# ==========================================
# MODULE: XUẤT EXCEL HỒ SƠ CHI ĐOÀN
# ==========================================
@app.route('/export_class_dashboard/<int:branch_id>')
def export_class_dashboard(branch_id):
    try:
        with session_scope() as db_session:
            active_year = db_session.query(SchoolYear).filter_by(is_active=True).first()
            if not active_year: return redirect(url_for('class_dashboard'))

            branch = db_session.query(Branch).filter_by(id=branch_id).first()
            if not branch: return redirect(url_for('class_dashboard'))

            group_val = branch.group or "Nhóm 1"
            weekly_scores = []
            weekly_scores_db = db_session.query(WeeklyScore).filter_by(branch_id=branch.id).order_by(WeeklyScore.id).all()
            for sc in weekly_scores_db:
                all_in_week = db_session.query(WeeklyScore).join(Branch).filter(WeeklyScore.week == sc.week, Branch.school_year_id == active_year.id).all()
                same_group_scores = [s for s in all_in_week if (s.branch.group or "Nhóm 1") == group_val]
                same_group_scores.sort(key=lambda x: float(x.total_score or 0), reverse=True)
                rk = 1
                for i, s in enumerate(same_group_scores):
                    if i > 0 and float(s.total_score or 0) < float(same_group_scores[i-1].total_score or 0): rk = i + 1
                    if s.branch_id == branch.id: break
                
                # [ĐÃ NÂNG CẤP LẠI]: Tính Tổng số con điểm Tốt (Không nhân hệ số)
                so_luong_diem_tot = int(sc.count_9 or 0) + int(sc.count_10 or 0)
                if "2" in str(group_val): so_luong_diem_tot += int(sc.count_8 or 0)
                
                weekly_scores.append({
                    'week': sc.week, 'rating': sc.week_rating or "-", 'score_truc': sc.score_truc or 100,
                    'diem_tot': so_luong_diem_tot, 'note': sc.note or "", 'total_score': sc.total_score or 0, 'rank': rk
                })
            
            red_star_ids = [rs.id for rs in db_session.query(RedStar).filter_by(branch_id=branch.id).all()]
            assignments = db_session.query(Assignment).filter(Assignment.red_star_id.in_(red_star_ids)).order_by(Assignment.week_number.desc(), Assignment.shift).all() if red_star_ids else []

            import json, os
            base_dir = os.path.dirname(os.path.abspath(__file__))
            zones_map = {}
            monitoring_assignments = []
            if os.path.exists(os.path.join(base_dir, "config", "class_zones.json")):
                with open(os.path.join(base_dir, "config", "class_zones.json"), "r", encoding="utf-8") as f:
                    try: zones_map = json.load(f)
                    except: pass
            target_zone_names = [zone for zone, classes in zones_map.items() if str(branch.name or '').strip().upper() in [str(c).strip().upper() for c in classes]] if zones_map else []
            if target_zone_names:
                monitoring_assignments = db_session.query(Assignment).join(DutyArea).join(RedStar).join(Branch).filter(
                    DutyArea.name.in_(target_zone_names), Branch.school_year_id == active_year.id
                ).order_by(Assignment.week_number.desc(), Assignment.shift).all()

            import openpyxl
            from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
            import io
            from flask import send_file

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = f"Ho_So_{branch.name}"

            font_title = Font(name="Times New Roman", size=14, bold=True)
            font_header = Font(name="Times New Roman", size=12, bold=True, color="FFFFFF")
            font_normal = Font(name="Times New Roman", size=12)
            align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
            border_thin = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")

            ws.merge_cells('A1:G1')
            ws['A1'] = f"HỒ SƠ TRA CỨU CHI ĐOÀN {branch.name.upper()}"
            ws['A1'].font = font_title; ws['A1'].alignment = align_center

            ws.merge_cells('A2:G2')
            ws['A2'] = f"GVCN: {branch.gvcn or '....................'} | Sĩ số: {branch.si_so} | Nhóm: {branch.group or '1'}"
            ws['A2'].font = Font(name="Times New Roman", size=12, italic=True); ws['A2'].alignment = align_center

            current_row = 4

            ws.merge_cells(f'A{current_row}:G{current_row}')
            ws[f'A{current_row}'] = "I. KẾT QUẢ THI ĐUA TỪNG TUẦN"
            ws[f'A{current_row}'].font = Font(name="Times New Roman", size=13, bold=True, color="C00000")
            current_row += 1

            # Đã đổi Header: Điểm Tốt -> SL Điểm Tốt
            headers_1 = ["Tuần", "Tổng điểm", "Hạng", "Xếp loại", "SL Điểm Tốt", "Ghi chú Vi phạm", ""]
            for col_num, h_text in enumerate(headers_1, 1):
                c = ws.cell(row=current_row, column=col_num, value=h_text)
                c.font = font_header; c.alignment = align_center; c.border = border_thin; c.fill = fill_header
            ws.cell(row=current_row, column=7).border = border_thin
            ws.cell(row=current_row, column=7).fill = fill_header
            ws.merge_cells(start_row=current_row, start_column=6, end_row=current_row, end_column=7)
            current_row += 1

            if weekly_scores:
                for sc in weekly_scores:
                    row_data = [sc['week'], sc['total_score'], f"Hạng {sc['rank']}", sc['rating'], sc['diem_tot'], sc['note']]
                    for col_num, val in enumerate(row_data, 1):
                        c = ws.cell(row=current_row, column=col_num, value=val)
                        c.font = font_normal; c.border = border_thin
                        c.alignment = Alignment(horizontal="left", vertical="center") if col_num == 6 else align_center
                    ws.cell(row=current_row, column=7).border = border_thin
                    ws.merge_cells(start_row=current_row, start_column=6, end_row=current_row, end_column=7)
                    current_row += 1
            else:
                ws.merge_cells(f'A{current_row}:G{current_row}')
                ws[f'A{current_row}'] = "Chưa có dữ liệu thi đua"
                ws[f'A{current_row}'].font = font_normal; ws[f'A{current_row}'].alignment = align_center; ws[f'A{current_row}'].border = border_thin
                current_row += 1

            current_row += 2

            ws.merge_cells(f'A{current_row}:G{current_row}')
            ws[f'A{current_row}'] = "II. DANH SÁCH SAO ĐỎ ĐÃ TRỰC CHẤM ĐIỂM LỚP"
            ws[f'A{current_row}'].font = Font(name="Times New Roman", size=13, bold=True, color="C00000")
            current_row += 1

            headers_3 = ["Tuần", "Ca trực", "Học sinh trực", "Thuộc Chi đoàn", "Khu vực trực", "Đánh giá KQ", "Ghi chú"]
            for col_num, h_text in enumerate(headers_3, 1):
                c = ws.cell(row=current_row, column=col_num, value=h_text)
                c.font = font_header; c.alignment = align_center; c.border = border_thin; c.fill = fill_header
            current_row += 1

            if monitoring_assignments:
                for assign in monitoring_assignments:
                    row_data = [f"Tuần {assign.week_number}", assign.shift, assign.red_star.full_name if assign.red_star else "-", 
                                assign.red_star.branch.name if assign.red_star and assign.red_star.branch else "-", 
                                assign.duty_area.name if assign.duty_area else "-", "", ""]
                    for col_num, val in enumerate(row_data, 1):
                        c = ws.cell(row=current_row, column=col_num, value=val)
                        c.font = font_normal; c.border = border_thin; c.alignment = align_center
                    current_row += 1
            else:
                ws.merge_cells(f'A{current_row}:G{current_row}')
                ws[f'A{current_row}'] = "Không có thông tin"
                ws[f'A{current_row}'].font = font_normal; ws[f'A{current_row}'].alignment = align_center; ws[f'A{current_row}'].border = border_thin
                current_row += 1

            current_row += 2

            ws.merge_cells(f'A{current_row}:G{current_row}')
            ws[f'A{current_row}'] = "III. LỊCH PHÂN CÔNG TRỰC CỦA HỌC SINH LỚP NÀY"
            ws[f'A{current_row}'].font = Font(name="Times New Roman", size=13, bold=True, color="C00000")
            current_row += 1

            headers_2 = ["Tuần", "Ca trực", "Học sinh trực", "Khu vực trực", "Nhiệm vụ", "Đánh giá KQ", "Ghi chú"]
            for col_num, h_text in enumerate(headers_2, 1):
                c = ws.cell(row=current_row, column=col_num, value=h_text)
                c.font = font_header; c.alignment = align_center; c.border = border_thin; c.fill = fill_header
            current_row += 1

            if assignments:
                for assign in assignments:
                    row_data = [f"Tuần {assign.week_number}", assign.shift, assign.red_star.full_name if assign.red_star else "-", 
                                assign.duty_area.name if assign.duty_area else "-", "", "", ""]
                    for col_num, val in enumerate(row_data, 1):
                        c = ws.cell(row=current_row, column=col_num, value=val)
                        c.font = font_normal; c.border = border_thin; c.alignment = align_center
                    current_row += 1
            else:
                ws.merge_cells(f'A{current_row}:G{current_row}')
                ws[f'A{current_row}'] = "Không có lịch phân công"
                ws[f'A{current_row}'].font = font_normal; ws[f'A{current_row}'].alignment = align_center; ws[f'A{current_row}'].border = border_thin
                current_row += 1

            ws.column_dimensions['A'].width = 10; ws.column_dimensions['B'].width = 12; ws.column_dimensions['C'].width = 12
            ws.column_dimensions['D'].width = 16; ws.column_dimensions['E'].width = 15; ws.column_dimensions['F'].width = 15
            ws.column_dimensions['G'].width = 30

            log_system_action("XUẤT EXCEL", f"Xuất Hồ sơ tra cứu Chi đoàn {branch.name}")
            out = io.BytesIO(); wb.save(out); out.seek(0)
            return send_file(out, download_name=f"Ho_So_{branch.name.replace(' ', '_')}.xlsx", as_attachment=True)

    except Exception as e:
        import traceback
        traceback.print_exc() 
        flash(f"Lỗi xuất Excel: {e}", "error")
        return redirect(url_for('class_dashboard'))
    
# ==========================================
# API: ĐỌC / GHI ĐIỂM GỐC TỪNG MÔN HỌC (RAWSCORE)
# ==========================================
@app.route('/api/raw_scores/<week_name>/<int:branch_id>', methods=['GET', 'POST'])
def handle_raw_scores(week_name, branch_id):
    try:
        from database.database import session_scope
        from database.models import Branch, SchoolYear, RawScore
        with session_scope() as db_session:
            branch = db_session.query(Branch).filter_by(id=branch_id).first()
            if not branch: 
                return {"error": "Lớp không tồn tại"}, 404
            
            active_year = db_session.query(SchoolYear).filter_by(is_active=True).first()
            year_id = active_year.id if active_year else 0
            
            # [CHÌA KHÓA VÀNG]: Gắn chặt Tên Tuần với ID Năm Học (VD: Tuần 1_Y2) để không bao giờ bị lẫn lộn giữa các năm
            safe_week_key = f"{week_name}_Y{year_id}"
            
            if request.method == 'GET':
                # Đổi week_name thành safe_week_key
                records = db_session.query(RawScore).filter_by(week=safe_week_key, branch_name=branch.name).all()
                data = [{"subj": r.subject, "c10": r.c10, "c9": r.c9, "c8": r.c8} for r in records]
                return {"data": data}

            if request.method == 'POST':
                raw_list = request.json.get('raw_list', [])
                # Đổi week_name thành safe_week_key
                db_session.query(RawScore).filter_by(week=safe_week_key, branch_name=branch.name).delete()
                
                for item in raw_list:
                    subj = str(item.get("subj", "")).strip()
                    if not subj or subj == "Điểm đã nhập": continue
                    rs = RawScore(
                        week=safe_week_key, # Đổi week_name thành safe_week_key
                        branch_name=branch.name,
                        subject=subj,
                        c10=int(item.get("c10", 0)),
                        c9=int(item.get("c9", 0)),
                        c8=int(item.get("c8", 0))
                    )
                    db_session.add(rs)
                return {"status": "success"}
    except Exception as e:
        return {"error": str(e)}, 500

# ==========================================
# MODULE: XEM TRƯỚC EXCEL BÁO CÁO TUẦN
# ==========================================
@app.route('/preview_report')
def preview_report():
    try:
        with session_scope() as db_session:
            week_name = request.args.get('week', 'Tuần 1')
            active_year = db_session.query(SchoolYear).filter(SchoolYear.is_active == True).first()
            
            if not active_year:
                flash("Chưa có năm học kích hoạt!", "error")
                return redirect(url_for('weekly'))
            
            current_scores = db_session.query(WeeklyScore).join(Branch).filter(
                WeeklyScore.week == week_name, Branch.school_year_id == active_year.id
            ).all()
            
            try:
                week_num = int(week_name.replace("Tuần ", "").strip())
                prev_week_name = f"Tuần {week_num - 1}"
            except:
                prev_week_name = None
                
            prev_rank_map = {}
            if prev_week_name:
                prev_scores = db_session.query(WeeklyScore).join(Branch).filter(
                    WeeklyScore.week == prev_week_name, Branch.school_year_id == active_year.id
                ).all()
                
                prev_data = {}
                for sc in prev_scores:
                    grp = sc.branch.group or "Nhóm 1"
                    if grp not in prev_data: prev_data[grp] = []
                    prev_data[grp].append(sc)
                for grp, lst in prev_data.items():
                    lst.sort(key=lambda x: float(x.total_score or 0), reverse=True)
                    rk = 1
                    for i, s in enumerate(lst):
                        if i > 0 and float(s.total_score or 0) < float(lst[i-1].total_score or 0):
                            rk = i + 1
                        prev_rank_map[s.branch_id] = rk

            report_data = {}
            start_date_str = ""
            end_date_str = ""
            
            for sc in current_scores:
                if sc.start_date: start_date_str = sc.start_date
                if sc.end_date: end_date_str = sc.end_date
                
                b = sc.branch
                grp = b.group or "Nhóm 1"
                if grp not in report_data: report_data[grp] = []
                
                so_luong_diem_tot = int(sc.count_9 or 0) + int(sc.count_10 or 0)
                if "2" in str(grp): 
                    so_luong_diem_tot += int(sc.count_8 or 0)
                
                report_data[grp].append({
                    'branch_id': b.id,
                    'branch_name': b.name,
                    'total_score': float(sc.total_score or 0),
                    'diem_tot': so_luong_diem_tot,
                    'note': sc.note or "",
                    'prev_rank': prev_rank_map.get(b.id, "N/A"),
                    'current_rank': 0
                })
                
            for grp, lst in report_data.items():
                lst.sort(key=lambda x: x['total_score'], reverse=True)
                rk = 1
                for i, item in enumerate(lst):
                    if i > 0 and item['total_score'] < lst[i-1]['total_score']: rk = i + 1
                    item['current_rank'] = rk

            if start_date_str and "-" in start_date_str:
                try: 
                    p = start_date_str.split('-')
                    start_date_str = f"{p[2]}/{p[1]}/{p[0]}"
                except: pass
            if end_date_str and "-" in end_date_str:
                try:
                    p = end_date_str.split('-')
                    end_date_str = f"{p[2]}/{p[1]}/{p[0]}"
                except: pass

            # [BẢN VÁ]: Nếu CSDL chưa lưu ngày, tự động gắn dấu chấm để giữ khung văn bản
            if not start_date_str: start_date_str = "......"
            if not end_date_str: end_date_str = "......"

            return render_template('preview_report.html', 
                                   report_data=report_data, 
                                   week_name=week_name, 
                                   start_date=start_date_str,
                                   end_date=end_date_str)
    except Exception as e:
        flash(f"Lỗi xem trước báo cáo: {e}", "error")
        return redirect(url_for('weekly'))

@app.route('/export_weekly_excel')
def export_weekly_excel():
    try:
        with session_scope() as db_session:
            week_name = request.args.get('week', 'Tuần 1')
            active_year = db_session.query(SchoolYear).filter(SchoolYear.is_active == True).first()
            if not active_year: return redirect(url_for('weekly'))
            
            current_scores = db_session.query(WeeklyScore).join(Branch).filter(WeeklyScore.week == week_name, Branch.school_year_id == active_year.id).all()
            try: week_num = int(week_name.replace("Tuần ", "").strip()); prev_week_name = f"Tuần {week_num - 1}"
            except: prev_week_name = None
                
            prev_rank_map = {}
            if prev_week_name:
                prev_scores = db_session.query(WeeklyScore).join(Branch).filter(WeeklyScore.week == prev_week_name, Branch.school_year_id == active_year.id).all()
                prev_data = {}
                for sc in prev_scores:
                    grp = sc.branch.group or "Nhóm 1"
                    if grp not in prev_data: prev_data[grp] = []
                    prev_data[grp].append(sc)
                    
                for grp, lst in prev_data.items():
                    # Xếp hạng đa tầng cho tuần trước
                    lst.sort(key=lambda x: (
                        -float(x.total_score or 0),
                        float(x.score_tru or 0),
                        -float(x.score_cong or 0),
                        x.branch.name
                    ))
                    rk = 1
                    for i, s in enumerate(lst):
                        if i > 0:
                            prev = lst[i-1]
                            if not (float(s.total_score or 0) == float(prev.total_score or 0) and 
                                    float(s.score_tru or 0) == float(prev.score_tru or 0) and 
                                    float(s.score_cong or 0) == float(prev.score_cong or 0)):
                                rk = i + 1
                        prev_rank_map[s.branch_id] = rk

            report_data = {}; start_date_str = ""; end_date_str = ""
            for sc in current_scores:
                if sc.start_date: start_date_str = sc.start_date
                if sc.end_date: end_date_str = sc.end_date
                b = sc.branch; grp = b.group or "Nhóm 1"
                if grp not in report_data: report_data[grp] = []
                
                so_luong_diem_tot = int(sc.count_9 or 0) + int(sc.count_10 or 0)
                if "2" in str(grp): 
                    so_luong_diem_tot += int(sc.count_8 or 0)
                    
                report_data[grp].append({
                    'branch_name': b.name, 'total_score': float(sc.total_score or 0),
                    'score_tru': float(sc.score_tru or 0), 'score_cong': float(sc.score_cong or 0),
                    'diem_tot': so_luong_diem_tot,
                    'note': sc.note or "", 'prev_rank': prev_rank_map.get(b.id, "N/A")
                })
                
            if start_date_str and "-" in start_date_str:
                try: 
                    p = start_date_str.split('-')
                    start_date_str = f"{p[2]}/{p[1]}/{p[0]}"
                except: pass
            if end_date_str and "-" in end_date_str:
                try:
                    p = end_date_str.split('-')
                    end_date_str = f"{p[2]}/{p[1]}/{p[0]}"
                except: pass

            # [BẢN VÁ]: Gắn mặc định dấu chấm nếu CSDL không có dữ liệu ngày
            if not start_date_str: start_date_str = "......"
            if not end_date_str: end_date_str = "......"
                
            import openpyxl
            from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
            from openpyxl.worksheet.page import PageMargins
            import io
            from flask import send_file

            wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Thi Đua Tuần"
            ws.page_setup.paperSize = ws.PAPERSIZE_A4; ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT 
            ws.page_margins = PageMargins(left=0.5, right=0.5, top=0.5, bottom=0.5, header=0.5, footer=0.5)
            
            font_title = Font(name="Times New Roman", size=11, bold=True)
            font_main_header = Font(name="Times New Roman", size=14, bold=True)
            font_data = Font(name="Times New Roman", size=11)
            align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
            align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
            fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
            font_header_color = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
            thin_border = Border(left=Side(style='thin', color='000000'), right=Side(style='thin', color='000000'), top=Side(style='thin', color='000000'), bottom=Side(style='thin', color='000000'))
            
            ws.merge_cells("A1:D1"); ws.cell(row=1, column=1, value="ĐOÀN TRƯỜNG THPT THANH HÒA").font = font_title; ws.cell(row=1, column=1).alignment = align_center
            ws.merge_cells("E1:G1"); ws.cell(row=1, column=5, value="ĐOÀN TNCS HỒ CHÍ MINH").font = font_title; ws.cell(row=1, column=5).alignment = align_center
            ws.merge_cells("A2:G2"); ws["A2"] = f"BẢNG TỔNG HỢP KẾT QUẢ THI ĐUA - {week_name.upper()}"; ws["A2"].font = font_main_header; ws["A2"].alignment = align_center
            
            # [BẢN VÁ]: Luôn ghi ra dòng Thời gian, không dùng lệnh IF ẩn đi nữa
            ws.merge_cells("A3:G3")
            ws["A3"] = f"(Từ ngày {start_date_str} đến ngày {end_date_str})"
            ws["A3"].font = Font(name="Times New Roman", size=12, italic=True)
            ws["A3"].alignment = align_center
            
            current_row = 5
            for group in sorted(report_data.keys()):
                group_items = report_data[group]
                # Xếp hạng đa tầng khi in ra Excel
                group_items.sort(key=lambda x: (
                    -float(x['total_score']),
                    float(x['score_tru']),
                    -float(x['score_cong']),
                    x['branch_name']
                ))
                
                curr_rank = 1
                for i, item in enumerate(group_items):
                    if i > 0:
                        prev = group_items[i-1]
                        if not (item['total_score'] == prev['total_score'] and 
                                item['score_tru'] == prev['score_tru'] and 
                                item['score_cong'] == prev['score_cong']):
                            curr_rank = i + 1
                    item['current_rank'] = curr_rank

                ws.cell(row=current_row, column=1, value=str(group)).font = Font(name="Times New Roman", size=11, bold=True)
                current_row += 1
                
                headers = ["STT", "Lớp", "X.H Tuần Trước", "Hạng Hiện Tại", "Tổng điểm", "SL Điểm Tốt", "Ghi chú"]
                for col_idx, h_text in enumerate(headers, 1):
                    cell = ws.cell(row=current_row, column=col_idx, value=h_text)
                    cell.font = font_header_color; cell.fill = fill_header; cell.alignment = align_center; cell.border = thin_border
                current_row += 1
                
                for idx, row_data in enumerate(group_items, 1):
                    val_tot = int(row_data["total_score"]) if row_data["total_score"].is_integer() else row_data["total_score"]
                    row_values = [idx, row_data["branch_name"], row_data["prev_rank"], row_data["current_rank"], val_tot, row_data["diem_tot"], row_data["note"]]
                    for col_idx, val in enumerate(row_values, 1):
                        cell = ws.cell(row=current_row, column=col_idx, value=val)
                        cell.font = font_data; cell.border = thin_border
                        if col_idx <= 6: cell.alignment = align_center
                        else: cell.alignment = align_left
                    current_row += 1
                current_row += 1 
            
            current_row += 1
            ws.merge_cells(start_row=current_row, start_column=5, end_row=current_row, end_column=7)
            ws.cell(row=current_row, column=5, value="TM/BCH ĐOÀN TRƯỜNG").font = font_title; ws.cell(row=current_row, column=5).alignment = align_center
            
            for col_letter, width in {'A': 5.5, 'B': 8.5, 'C': 11.5, 'D': 11.5, 'E': 11.5, 'F': 11.0, 'G': 30.0}.items():
                ws.column_dimensions[col_letter].width = width
            
            log_system_action("XUẤT EXCEL", f"Xuất Báo cáo Điểm {week_name}")
            out = io.BytesIO()
            wb.save(out)
            out.seek(0)
            return send_file(out, download_name=f"Bao_Cao_Thi_Dua_{week_name.replace(' ', '_')}.xlsx", as_attachment=True)
    except Exception as e:
        flash(f"Lỗi xuất file Excel: {str(e)}", "error")
        return redirect(url_for('weekly'))
# ==========================================
# MODULE: XUẤT EXCEL THI ĐUA THÁNG
# ==========================================
@app.route('/export_monthly_excel', methods=['POST'])
def export_monthly_excel():
    try:
        with session_scope() as db_session:
            active_year = db_session.query(SchoolYear).filter_by(is_active=True).first()
            if not active_year:
                flash("Chưa có năm học nào được kích hoạt!", "error")
                return redirect(url_for('monthly'))

            selected_month = request.form.get('month', 'Tháng...')
            selected_weeks = request.form.getlist('weeks')

            if not selected_weeks:
                flash("Vui lòng chọn ít nhất 1 tuần để xuất báo cáo!", "error")
                return redirect(url_for('monthly'))

            groups_data = {}
            branches = db_session.query(Branch).filter(Branch.school_year_id == active_year.id).all()
            
            for b in branches:
                grp = b.group or "Nhóm 1"
                if grp not in groups_data:
                    groups_data[grp] = []
                    
                scores = db_session.query(WeeklyScore).filter(
                    WeeklyScore.branch_id == b.id,
                    WeeklyScore.week.in_(selected_weeks)
                ).all()
                
                week_scores = {s.week: (s.total_score or 0.0) for s in scores}
                total_score = sum(week_scores.values())
                
                groups_data[grp].append({
                    'branch_name': b.name,
                    'week_scores': week_scores,
                    'total_score': total_score,
                    'gvcn': b.gvcn
                })
            
            for grp, lst in groups_data.items():
                lst.sort(key=lambda x: x['total_score'], reverse=True)
                rk = 1
                for i, d in enumerate(lst):
                    if i > 0 and d['total_score'] < lst[i-1]['total_score']:
                        rk = i + 1
                    d['rank'] = rk

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Báo Cáo Thi Đua"

            font_title = Font(name="Times New Roman", size=12, bold=True)
            font_main_header = Font(name="Times New Roman", size=16, bold=True)
            font_header_table = Font(name="Times New Roman", size=12, bold=True)
            font_data = Font(name="Times New Roman", size=12)
            align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
            align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
            border_thin = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

            num_cols = 2 + len(selected_weeks) + 3 
            last_col_letter = get_column_letter(num_cols)

            ws.merge_cells("A1:C1")
            ws["A1"] = "ĐOÀN TRƯỜNG THPT THANH HÒA"
            ws["A1"].font = font_title
            ws["A1"].alignment = align_center

            ws.merge_cells(f"E1:{last_col_letter}1" if num_cols >= 5 else f"D1:{last_col_letter}1")
            ws["E1" if num_cols >= 5 else "D1"] = "ĐOÀN TNCS HỒ CHÍ MINH"
            ws["E1" if num_cols >= 5 else "D1"].font = font_title
            ws["E1" if num_cols >= 5 else "D1"].alignment = align_center

            ws.merge_cells(f"A3:{last_col_letter}3")
            ws["A3"] = f"ĐIỂM THI ĐUA {selected_month.upper()}"
            ws["A3"].font = font_main_header
            ws["A3"].alignment = align_center
            
            week_str = ", ".join(selected_weeks)
            ws.merge_cells(f"A4:{last_col_letter}4")
            ws["A4"] = f"(Tổng hợp điểm từ các tuần: {week_str})"
            ws["A4"].font = Font(name="Times New Roman", size=12, italic=True)
            ws["A4"].alignment = align_center
            
            mid_col = num_cols // 2 if num_cols > 3 else 2
            ws.merge_cells(f"{get_column_letter(mid_col)}5:{get_column_letter(mid_col+2)}5")
            ws[f"{get_column_letter(mid_col)}5"] = f"NĂM HỌC: {active_year.name}"
            ws[f"{get_column_letter(mid_col)}5"].font = Font(name="Times New Roman", size=12, bold=True)
            ws[f"{get_column_letter(mid_col)}5"].alignment = align_center

            current_row = 6
            for grp in sorted(groups_data.keys()):
                ws.cell(row=current_row, column=1, value=str(grp)).font = Font(name="Times New Roman", size=12, bold=True, italic=True)
                current_row += 1
                
                headers = ["STT", "LỚP"] + [w.replace("Tuần ", "T.") for w in selected_weeks] + ["ĐIỂM", "HẠNG\nHIỆN TẠI", "GVCN"]
                for col_idx, h_text in enumerate(headers, 1):
                    cell = ws.cell(row=current_row, column=col_idx, value=h_text)
                    cell.font = font_header_table
                    cell.alignment = align_center
                    cell.border = border_thin
                current_row += 1
                
                for idx, item in enumerate(groups_data[grp], 1):
                    val_tot = int(item['total_score']) if float(item['total_score']).is_integer() else item['total_score']
                    row_vals = [idx, item['branch_name']]
                    for w in selected_weeks:
                        ws_score = item['week_scores'].get(w, 0.0)
                        row_vals.append(int(ws_score) if float(ws_score).is_integer() else ws_score)
                    row_vals.extend([val_tot, item['rank'], item['gvcn']])
                    
                    for col_idx, val in enumerate(row_vals, 1):
                        cell = ws.cell(row=current_row, column=col_idx, value=val)
                        cell.font = font_data
                        cell.border = border_thin
                        if col_idx == 2 or col_idx == num_cols:
                            cell.alignment = align_center if col_idx == 2 else align_left
                        else:
                            cell.alignment = align_center
                    current_row += 1
                current_row += 1 

            current_row += 1
            start_col = num_cols - 2 if num_cols >= 3 else num_cols
            ws.merge_cells(start_row=current_row, start_column=start_col-1, end_row=current_row, end_column=num_cols)
            sign_cell = ws.cell(row=current_row, column=start_col-1, value="TM. BCH ĐOÀN TRƯỜNG")
            sign_cell.font = font_title
            sign_cell.alignment = align_center

            ws.column_dimensions['A'].width = 6
            ws.column_dimensions['B'].width = 10
            for i in range(len(selected_weeks)):
                ws.column_dimensions[get_column_letter(3+i)].width = 7
            ws.column_dimensions[get_column_letter(num_cols - 2)].width = 10 
            ws.column_dimensions[get_column_letter(num_cols - 1)].width = 12 
            ws.column_dimensions[get_column_letter(num_cols)].width = 25 

            log_system_action("XUẤT EXCEL", f"Xuất Báo cáo Điểm {selected_month}")
            out = io.BytesIO()
            wb.save(out)
            out.seek(0)
            return send_file(out, download_name=f"Bao_Cao_Diem_{selected_month.replace(' ', '_')}.xlsx", as_attachment=True)

    except Exception as e:
        import traceback
        traceback.print_exc()
        flash(f"Lỗi xuất Excel: {str(e)}", "error")
        return redirect(url_for('monthly'))

# ==========================================
# MODULE: THI ĐUA HỌC KỲ (LẤY DỮ LIỆU TỪ THÁNG & LƯU HỆ THỐNG)
# ==========================================
@app.route('/semester', methods=['GET', 'POST'])
def semester():
    try:
        with session_scope() as db_session:
            active_year = db_session.query(SchoolYear).filter_by(is_active=True).first()
            
            available_months = []
            used_by_other_semesters = set()
            current_semester_months = set()
            semester_data = {}
            
            selected_semester = request.form.get('semester') or request.args.get('semester') or 'Học kỳ 1'
            selected_months = request.form.getlist('months')
            action = request.form.get('action', 'view')

            if active_year:
                months_db = db_session.query(MonthlyRecord.month_name).filter(
                    MonthlyRecord.school_year_id == active_year.id,
                    MonthlyRecord.month_name.like('Tháng%')
                ).distinct().all()
                
                school_order = ["Tháng 9", "Tháng 10", "Tháng 11", "Tháng 12", "Tháng 1", "Tháng 2", "Tháng 3", "Tháng 4", "Tháng 5"]
                raw_months = [m[0] for m in months_db if m[0]]
                available_months = sorted(raw_months, key=lambda x: school_order.index(x) if x in school_order else 99)

                all_sem_records = db_session.query(MonthlyRecord).filter(
                    MonthlyRecord.school_year_id == active_year.id,
                    MonthlyRecord.month_name.like("Học kỳ%")
                ).all()
                
                for rec in all_sem_records:
                    if rec.weeks_used:
                        m_list = [m.strip() for m in rec.weeks_used.split(",") if m.strip()]
                        if rec.month_name == selected_semester:
                            current_semester_months.update(m_list)
                        else:
                            used_by_other_semesters.update(m_list)

                if request.method == 'GET' and not selected_months:
                    selected_months = list(current_semester_months)
                    selected_months = sorted(selected_months, key=lambda x: school_order.index(x) if x in school_order else 99)

                if selected_months:
                    prev_sem_name = "Học kỳ 1" if selected_semester == "Học kỳ 2" else None
                    prev_ranks = {}
                    if prev_sem_name:
                        p_records = db_session.query(MonthlyRecord).filter(
                            MonthlyRecord.school_year_id == active_year.id,
                            MonthlyRecord.month_name == prev_sem_name
                        ).all()
                        for pr in p_records:
                            prev_ranks[pr.branch_id] = pr.rank

                    branches = db_session.query(Branch).filter(Branch.school_year_id == active_year.id).all()
                    for b in branches:
                        grp = b.group or "Nhóm 1"
                        if grp not in semester_data: semester_data[grp] = []
                        
                        m_scores = db_session.query(MonthlyRecord).filter(
                            MonthlyRecord.branch_id == b.id,
                            MonthlyRecord.month_name.in_(selected_months)
                        ).all()
                        
                        month_scores_dict = {s.month_name: (s.total_score or 0.0) for s in m_scores}
                        total_score = sum(month_scores_dict.values())
                        
                        semester_data[grp].append({
                            'branch_id': b.id,
                            'branch_name': b.name,
                            'gvcn': b.gvcn,
                            'month_scores': month_scores_dict,
                            'total_score': total_score,
                            'prev_rank': prev_ranks.get(b.id, "-")
                        })
                        
                    for grp, lst in semester_data.items():
                        # Sắp xếp theo Tổng điểm giảm dần, nếu bằng điểm thì xếp A-Z
                        lst.sort(key=lambda x: (-float(x['total_score']), x['branch_name']))
                        rk = 1
                        for i, d in enumerate(lst):
                            if i > 0:
                                prev = lst[i-1]
                                if float(d['total_score']) != float(prev['total_score']):
                                    rk = i + 1
                            d['rank'] = rk
                            
                            p_rk = d['prev_rank']
                            diff_val = 0
                            change_str = "-"
                            if p_rk != "-" and isinstance(p_rk, int):
                                diff_val = p_rk - rk
                                if diff_val > 0: change_str = f"▲ Tăng {diff_val}"
                                elif diff_val < 0: change_str = f"▼ Giảm {abs(diff_val)}"
                                else: change_str = "▬ Giữ nguyên"
                            d['change_str'] = change_str
                            d['diff_val'] = diff_val

                    if action == 'save':
                        db_session.query(MonthlyRecord).filter(
                            MonthlyRecord.school_year_id == active_year.id,
                            MonthlyRecord.month_name == selected_semester
                        ).delete()
                        
                        months_str = ", ".join(selected_months)
                        for grp, lst in semester_data.items():
                            for d in lst:
                                db_session.add(MonthlyRecord(
                                    school_year_id=active_year.id,
                                    month_name=selected_semester,
                                    branch_id=d['branch_id'],
                                    total_score=d['total_score'],
                                    rank=d['rank'],
                                    weeks_used=months_str
                                ))
                        db_session.commit()
                        
                        log_system_action("LƯU ĐIỂM HỌC KỲ", f"Đã tính toán và chốt sổ điểm {selected_semester} (gộp từ: {months_str}).")
                        flash(f"✅ Đã lưu thành công dữ liệu {selected_semester} vào hệ thống!", "success")
                        return redirect(url_for('semester', semester=selected_semester))

            return render_template(
                'semester.html', 
                available_months=available_months, 
                used_by_other_semesters=list(used_by_other_semesters),
                semester_data=semester_data,
                selected_semester=selected_semester, 
                selected_months=selected_months
            )
    except Exception as e:
        import traceback
        traceback.print_exc()
        flash(f"Lỗi phân hệ thi đua học kỳ: {str(e)}", "error")
        return redirect(url_for('dashboard'))

# ==========================================
# MODULE: XUẤT EXCEL HỌC KỲ
# ==========================================
@app.route('/export_semester_excel', methods=['POST'])
def export_semester_excel():
    try:
        with session_scope() as db_session:
            active_year = db_session.query(SchoolYear).filter_by(is_active=True).first()
            if not active_year: return redirect(url_for('semester'))

            selected_semester = request.form.get('semester', 'Học kỳ 1')
            selected_months = request.form.getlist('months')
            is_hk1 = (selected_semester == "Học kỳ 1")

            if not selected_months:
                flash("Vui lòng chọn ít nhất 1 tháng để xuất Excel!", "error")
                return redirect(url_for('semester'))

            semester_data = {}
            prev_sem_name = "Học kỳ 1" if not is_hk1 else None
            prev_ranks = {}
            if prev_sem_name:
                p_records = db_session.query(MonthlyRecord).filter(MonthlyRecord.school_year_id == active_year.id, MonthlyRecord.month_name == prev_sem_name).all()
                for pr in p_records: prev_ranks[pr.branch_id] = pr.rank

            branches = db_session.query(Branch).filter(Branch.school_year_id == active_year.id).all()
            for b in branches:
                grp = b.group or "Nhóm 1"
                if grp not in semester_data: semester_data[grp] = []
                m_scores = db_session.query(MonthlyRecord).filter(MonthlyRecord.branch_id == b.id, MonthlyRecord.month_name.in_(selected_months)).all()
                
                month_scores_dict = {s.month_name: (s.total_score or 0.0) for s in m_scores}
                semester_data[grp].append({
                    'branch_name': b.name, 'gvcn': b.gvcn, 'month_scores': month_scores_dict,
                    'total_score': sum(month_scores_dict.values()), 'prev_rank': prev_ranks.get(b.id, "-")
                })
                
            for grp, lst in semester_data.items():
                lst.sort(key=lambda x: x['total_score'], reverse=True)
                rk = 1
                for i, d in enumerate(lst):
                    if i > 0 and d['total_score'] < lst[i-1]['total_score']: rk = i + 1
                    d['rank'] = rk
                    p_rk = d['prev_rank']
                    if p_rk != "-" and isinstance(p_rk, int):
                        diff_val = p_rk - rk
                        if diff_val > 0: d['change_str'] = f"▲ Tăng {diff_val}"
                        elif diff_val < 0: d['change_str'] = f"▼ Giảm {abs(diff_val)}"
                        else: d['change_str'] = "▬ Giữ nguyên"
                    else: d['change_str'] = "-"

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Báo Cáo Học Kỳ"
            ws.page_setup.paperSize = ws.PAPERSIZE_A4
            ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
            ws.page_margins = PageMargins(left=0.75, right=0.5, top=0.75, bottom=0.75, header=0.5, footer=0.5)

            font_bold = Font(name='Times New Roman', size=11, bold=True)
            font_normal = Font(name='Times New Roman', size=11)
            font_title = Font(name='Times New Roman', size=14, bold=True)
            align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
            align_left = Alignment(horizontal='left', vertical='center')
            border_thin = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            fill_header = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")

            month_headers = [m.replace("Tháng ", "T.") for m in selected_months]
            if is_hk1: base_headers = ["STT", "LỚP"] + month_headers + ["TỔNG ĐIỂM", "HẠNG\n HIỆN TẠI", "GVCN"]
            else: base_headers = ["STT", "LỚP"] + month_headers + ["TỔNG ĐIỂM", "XH. KỲ TRƯỚC", "HẠNG\n HIỆN TẠI", "TĂNG/GIẢM", "GVCN"]
            
            total_cols = len(base_headers)
            last_col_letter = get_column_letter(total_cols)

            ws.merge_cells('A1:C1')
            ws['A1'] = "ĐOÀN TRƯỜNG THPT THANH HÒA"
            ws['A1'].font = font_bold
            ws['A1'].alignment = align_left

            ws.merge_cells(f'E1:{last_col_letter}1' if total_cols >= 5 else f'D1:{last_col_letter}1')
            ws['E1' if total_cols >= 5 else 'D1'] = "ĐOÀN TNCS HỒ CHÍ MINH"
            ws['E1' if total_cols >= 5 else 'D1'].font = Font(name='Times New Roman', size=14, bold=True)
            ws['E1' if total_cols >= 5 else 'D1'].alignment = align_center

            ws.merge_cells(f'A3:{last_col_letter}3')
            ws['A3'] = f"ĐIỂM THI ĐUA {selected_semester.upper()}"
            ws['A3'].font = font_title
            ws['A3'].alignment = align_center
            ws.row_dimensions[3].height = 25.0
            
            mid_col = total_cols // 2 if total_cols > 3 else 2
            ws.merge_cells(f'{get_column_letter(mid_col)}4:{get_column_letter(mid_col+2)}4')
            ws[f'{get_column_letter(mid_col)}4'] = f"NĂM HỌC: {active_year.name}"
            ws[f'{get_column_letter(mid_col)}4'].font = font_bold
            ws[f'{get_column_letter(mid_col)}4'].alignment = align_center

            current_row = 6
            for grp in sorted(semester_data.keys()):
                ws.cell(row=current_row, column=1, value=str(grp)).font = Font(name='Times New Roman', size=12, bold=True, italic=True)
                current_row += 1

                for col, h in enumerate(base_headers, 1):
                    cell = ws.cell(row=current_row, column=col, value=h)
                    cell.font = font_bold
                    cell.alignment = align_center
                    cell.border = border_thin
                    cell.fill = fill_header
                current_row += 1

                for idx, r in enumerate(semester_data[grp], 1):
                    row_data = [idx, r['branch_name']]
                    for m in selected_months:
                        sc = r['month_scores'].get(m, 0.0)
                        row_data.append(int(sc) if float(sc).is_integer() else sc)
                    
                    tot = r['total_score']
                    row_data.append(int(tot) if float(tot).is_integer() else tot)

                    if is_hk1:
                        row_data.append(r['rank'])
                        row_data.append(r['gvcn'])
                    else:
                        row_data.append(r['prev_rank'])
                        row_data.append(r['rank'])
                        row_data.append(r['change_str'])
                        row_data.append(r['gvcn'])

                    for col, val in enumerate(row_data, 1):
                        cell = ws.cell(row=current_row, column=col, value=val)
                        cell.font = font_normal
                        cell.border = border_thin
                        cell.alignment = align_center if col != 2 and col != total_cols else align_left
                        
                        if not is_hk1 and base_headers[col-1] == "TĂNG/GIẢM":
                            if "Tăng" in str(val): cell.font = Font(name='Times New Roman', size=11, bold=True, color="008000")
                            elif "Giảm" in str(val): cell.font = Font(name='Times New Roman', size=11, bold=True, color="FF0000")

                    current_row += 1
                current_row += 1

            current_row += 1
            sign_col = total_cols - 1 if total_cols >= 3 else total_cols
            ws.merge_cells(start_row=current_row, start_column=sign_col-1, end_row=current_row, end_column=total_cols)
            sign_cell = ws.cell(row=current_row, column=sign_col-1, value="TM. BCH ĐOÀN TRƯỜNG")
            sign_cell.font = font_bold
            sign_cell.alignment = align_center

            ws.column_dimensions['A'].width = 5.5
            ws.column_dimensions['B'].width = 10.0
            for c in range(3, total_cols - 1):
                ws.column_dimensions[get_column_letter(c)].width = 7.0
            ws.column_dimensions[get_column_letter(total_cols - 2)].width = 10.0 
            ws.column_dimensions[get_column_letter(total_cols - 1)].width = 12.0 
            ws.column_dimensions[last_col_letter].width = 25.0 

            log_system_action("XUẤT EXCEL", f"Xuất Báo cáo Điểm {selected_semester}")
            out = io.BytesIO()
            wb.save(out)
            out.seek(0)
            return send_file(out, download_name=f"Bao_Cao_{selected_semester.replace(' ', '_')}.xlsx", as_attachment=True)

    except Exception as e:
        import traceback
        traceback.print_exc()
        flash(f"Lỗi xuất Excel: {str(e)}", "error")
        return redirect(url_for('semester'))

# ==========================================
# MODULE: TỔNG KẾT NĂM HỌC
# ==========================================
@app.route('/yearly', methods=['GET', 'POST'])
def yearly():
    try:
        with session_scope() as db_session:
            active_year = db_session.query(SchoolYear).filter_by(is_active=True).first()
            
            available_semesters = []
            yearly_data = {}
            
            selected_year_name = request.form.get('year_name') or request.args.get('year_name') or 'Năm học'
            if active_year and selected_year_name == 'Năm học':
                selected_year_name = f"Năm học {active_year.name}"
                
            selected_sems = request.form.getlist('semesters')
            action = request.form.get('action', 'view')

            if active_year:
                sems_db = db_session.query(MonthlyRecord.month_name).filter(
                    MonthlyRecord.school_year_id == active_year.id,
                    MonthlyRecord.month_name.like('Học kỳ%')
                ).distinct().all()
                
                available_semesters = sorted([m[0] for m in sems_db if m[0]])

                if request.method == 'GET' and not selected_sems:
                    selected_sems = available_semesters

                if selected_sems:
                    branches = db_session.query(Branch).filter(Branch.school_year_id == active_year.id).all()
                    for b in branches:
                        grp = b.group or "Nhóm 1"
                        if grp not in yearly_data: yearly_data[grp] = []
                        
                        sem_scores_records = db_session.query(MonthlyRecord).filter(
                            MonthlyRecord.branch_id == b.id,
                            MonthlyRecord.month_name.in_(selected_sems)
                        ).all()
                        
                        hk1_score = "-"
                        hk2_score = "-"
                        total_score = 0.0
                        
                        for s in sem_scores_records:
                            val = s.total_score or 0.0
                            if "1" in s.month_name: hk1_score = val
                            if "2" in s.month_name: hk2_score = val
                            total_score += val
                            
                        yearly_data[grp].append({
                            'branch_id': b.id,
                            'branch_name': b.name,
                            'gvcn': b.gvcn,
                            'hk1_score': hk1_score,
                            'hk2_score': hk2_score,
                            'total_score': total_score
                        })
                        
                    for grp, lst in yearly_data.items():
                        # Sắp xếp theo Tổng điểm giảm dần, nếu bằng điểm thì xếp A-Z
                        lst.sort(key=lambda x: (-float(x['total_score']), x['branch_name']))
                        rk = 1
                        for i, d in enumerate(lst):
                            if i > 0:
                                prev = lst[i-1]
                                if float(d['total_score']) != float(prev['total_score']):
                                    rk = i + 1
                            d['rank'] = rk

                    if action == 'save':
                        db_session.query(MonthlyRecord).filter(
                            MonthlyRecord.school_year_id == active_year.id,
                            MonthlyRecord.month_name == selected_year_name
                        ).delete()
                        
                        sems_str = ", ".join(selected_sems)
                        for grp, lst in yearly_data.items():
                            for d in lst:
                                db_session.add(MonthlyRecord(
                                    school_year_id=active_year.id,
                                    month_name=selected_year_name,
                                    branch_id=d['branch_id'],
                                    total_score=d['total_score'],
                                    rank=d['rank'],
                                    weeks_used=sems_str
                                ))
                        db_session.commit()
                        
                        log_system_action("LƯU ĐIỂM NĂM HỌC", f"Đã tính toán và chốt sổ điểm {selected_year_name} (gộp từ: {sems_str}).")
                        flash(f"✅ Đã lưu thành công dữ liệu Tổng kết {selected_year_name}!", "success")
                        return redirect(url_for('yearly', year_name=selected_year_name))

            return render_template(
                'yearly.html', 
                available_semesters=available_semesters, 
                yearly_data=yearly_data,
                selected_year_name=selected_year_name, 
                selected_sems=selected_sems,
                active_year=active_year
            )
    except Exception as e:
        import traceback
        traceback.print_exc()
        flash(f"Lỗi phân hệ Tổng kết Năm học: {str(e)}", "error")
        return redirect(url_for('dashboard'))

@app.route('/export_yearly_excel', methods=['POST'])
def export_yearly_excel():
    try:
        with session_scope() as db_session:
            active_year = db_session.query(SchoolYear).filter_by(is_active=True).first()
            if not active_year: return redirect(url_for('yearly'))

            selected_year_name = request.form.get('year_name', 'Năm học')
            selected_sems = request.form.getlist('semesters')

            if not selected_sems:
                flash("Vui lòng chọn ít nhất 1 học kỳ để xuất Excel!", "error")
                return redirect(url_for('yearly'))

            yearly_data = {}
            branches = db_session.query(Branch).filter(Branch.school_year_id == active_year.id).all()
            for b in branches:
                grp = b.group or "Nhóm 1"
                if grp not in yearly_data: yearly_data[grp] = []
                
                sem_scores = db_session.query(MonthlyRecord).filter(MonthlyRecord.branch_id == b.id, MonthlyRecord.month_name.in_(selected_sems)).all()
                hk1 = "-"; hk2 = "-"; tot = 0.0
                for s in sem_scores:
                    val = s.total_score or 0.0
                    if "1" in str(s.month_name): hk1 = val
                    if "2" in str(s.month_name): hk2 = val
                    tot += val
                    
                yearly_data[grp].append({'branch_name': b.name, 'gvcn': b.gvcn, 'hk1_score': hk1, 'hk2_score': hk2, 'total_score': tot})
                
            for grp, lst in yearly_data.items():
                lst.sort(key=lambda x: x['total_score'], reverse=True)
                rk = 1
                for i, d in enumerate(lst):
                    if i > 0 and d['total_score'] < lst[i-1]['total_score']: rk = i + 1
                    d['rank'] = rk

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Tổng Kết Năm Học"
            ws.page_setup.paperSize = ws.PAPERSIZE_A4
            ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
            ws.page_margins = PageMargins(left=0.75, right=0.5, top=0.75, bottom=0.75, header=0.5, footer=0.5)

            font_bold = Font(name='Times New Roman', size=11, bold=True)
            font_normal = Font(name='Times New Roman', size=11)
            font_title = Font(name='Times New Roman', size=14, bold=True)
            align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
            align_left = Alignment(horizontal='left', vertical='center')
            border_thin = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            fill_header = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")

            ws.merge_cells('A1:C1'); ws['A1'] = "ĐOÀN TRƯỜNG THPT THANH HÒA"; ws['A1'].font = font_bold; ws['A1'].alignment = align_left
            ws.merge_cells('E1:G1'); ws['E1'] = "ĐOÀN TNCS HỒ CHÍ MINH"; ws['E1'].font = Font(name='Times New Roman', size=14, bold=True); ws['E1'].alignment = align_center
            ws.merge_cells('A3:G3'); ws['A3'] = f"TỔNG KẾT THI ĐUA NĂM HỌC - {selected_year_name.upper()}"; ws['A3'].font = font_title; ws['A3'].alignment = align_center

            current_row = 5
            for grp in sorted(yearly_data.keys()):
                ws.cell(row=current_row, column=1, value=str(grp)).font = font_bold; current_row += 1

                headers = ["STT", "LỚP", "ĐIỂM HK1", "ĐIỂM HK2", "TỔNG ĐIỂM NĂM", "HẠNG HIỆN TẠI", "GVCN"]
                for col, h in enumerate(headers, 1):
                    cell = ws.cell(row=current_row, column=col, value=h)
                    cell.font = font_bold; cell.alignment = align_center; cell.border = border_thin; cell.fill = fill_header
                current_row += 1

                for idx, r in enumerate(yearly_data[grp], 1):
                    h1 = int(r['hk1_score']) if isinstance(r['hk1_score'], float) and r['hk1_score'].is_integer() else r['hk1_score']
                    h2 = int(r['hk2_score']) if isinstance(r['hk2_score'], float) and r['hk2_score'].is_integer() else r['hk2_score']
                    tt = int(r['total_score']) if isinstance(r['total_score'], float) and r['total_score'].is_integer() else r['total_score']
                    
                    row_data = [idx, r['branch_name'], h1, h2, tt, r['rank'], r['gvcn']]
                    for col, val in enumerate(row_data, 1):
                        cell = ws.cell(row=current_row, column=col, value=val)
                        cell.font = font_normal; cell.border = border_thin
                        cell.alignment = align_center if col in [1, 3, 4, 5, 6] else align_left
                    current_row += 1
                current_row += 1

            current_row += 1
            ws.merge_cells(start_row=current_row, start_column=5, end_row=current_row, end_column=7)
            ws.cell(row=current_row, column=5, value="TM. BCH ĐOÀN TRƯỜNG").font = font_bold; ws.cell(row=current_row, column=5).alignment = align_center

            widths = {'A': 5.5, 'B': 14.0, 'C': 12.0, 'D': 12.0, 'E': 15.0, 'F': 18.0, 'G': 25.0}
            for col, width in widths.items(): ws.column_dimensions[col].width = width

            log_system_action("XUẤT EXCEL", f"Xuất báo cáo Điểm {selected_year_name}")
            out = io.BytesIO(); wb.save(out); out.seek(0)
            return send_file(out, download_name=f"Tong_Ket_Nam_Hoc.xlsx", as_attachment=True)

    except Exception as e:
        import traceback
        traceback.print_exc()
        flash(f"Lỗi xuất Excel: {str(e)}", "error")
        return redirect(url_for('yearly'))

# =========================================================================
# THUẬT TOÁN ĐỌC DỮ LIỆU TỪ CƠ SỞ DỮ LIỆU SQLITE CHO BIỂU MẪU BÁO CÁO
# =========================================================================
def calculate_trimmed_good_points_web(session, week_name, branch_name, branch_group, max_mon, max_tot):
    """Tính lại số điểm tốt dựa trên dữ liệu thô từ CSDL và barem động hiện hành"""
    try:
        from database.models import SchoolYear, RawScore, WeeklyScore, Branch
        active_year = session.query(SchoolYear).filter_by(is_active=True).first()
        year_id = active_year.id if active_year else 0
        safe_week_key = f"{week_name}_Y{year_id}"
        
        records = session.query(RawScore).filter_by(week=safe_week_key, branch_name=branch_name).all()
        raw_scores = [{"c10": r.c10, "c9": r.c9, "c8": r.c8} for r in records]
    except Exception:
        raw_scores = []
        
    # [BẢN VÁ LỖI CỐT LÕI]: Nếu không quét file Excel, lấy điểm nhập tay từ Web làm dự phòng
    if not raw_scores: 
        branch = session.query(Branch).filter_by(name=branch_name, school_year_id=year_id).first()
        if branch:
            sc = session.query(WeeklyScore).filter_by(branch_id=branch.id, week=week_name).first()
            if sc:
                total_good = int(sc.count_10 or 0) + int(sc.count_9 or 0)
                if "2" in str(branch_group):
                    total_good += int(sc.count_8 or 0)
                return total_good 
        return 0
        
    f10, f9, f8 = 0, 0, 0
    for r in raw_scores:
        try:
            c10, c9, c8 = int(r.get('c10', 0)), int(r.get('c9', 0)), int(r.get('c8', 0))
        except ValueError: continue
        
        if "1" in str(branch_group): c8 = 0 
        
        k10 = min(c10, max_mon)
        k9 = min(c9, max_mon - k10)
        k8 = min(c8, max_mon - k10 - k9)
        f10 += k10; f9 += k9; f8 += k8
        
    total = f10 + f9 + f8
    return min(total, max_tot)

@app.route('/templates', methods=['GET', 'POST'])
def templates_report():
    try:
        with session_scope() as db_session:
            years = db_session.query(SchoolYear).order_by(SchoolYear.id.desc()).all()
            active_year = db_session.query(SchoolYear).filter_by(is_active=True).first()
            
            selected_year_name = request.args.get('year_name', active_year.name if active_year else '')
            report_type = request.args.get('report_type', 'Báo cáo Tuần')
            time_val = request.args.get('time_val', 'Tuần 1')
            group_val = request.args.get('group_val', 'Tất cả')

            selected_year = db_session.query(SchoolYear).filter_by(name=selected_year_name).first()
            
            time_options = []
            if "Tuần" in report_type:
                time_options = [f"Tuần {i}" for i in range(1, 39)]
            elif "Tháng" in report_type:
                time_options = ["Tháng 8", "Tháng 9", "Tháng 10", "Tháng 11", "Tháng 12", "Tháng 1", "Tháng 2", "Tháng 3", "Tháng 4", "Tháng 5"]
            elif "Học kỳ" in report_type:
                time_options = ["Học kỳ 1", "Học kỳ 2", "Cả năm học"]

            data_list = []
            start_date_str = "…"
            end_date_str = "…"

            if selected_year:
                max_tot, max_mon = 14, 4
                try:
                    settings = db_session.query(ScoreSettings).filter_by(school_year_id=selected_year.id).first()
                    if settings:
                        if hasattr(settings, 'max_diem_tot'): max_tot = int(settings.max_diem_tot)
                        if hasattr(settings, 'max_diem_mon'): max_mon = int(settings.max_diem_mon)
                except Exception: pass

                branch_query = db_session.query(Branch).filter_by(school_year_id=selected_year.id)
                if group_val != "Tất cả":
                    if group_val.startswith("Khối"):
                        khoi = group_val.replace("Khối", "").strip()
                        from sqlalchemy import or_
                        branch_query = branch_query.filter(or_(Branch.name.startswith(khoi), Branch.name.like(f"% {khoi}%")))
                    else:
                        branch_query = branch_query.filter_by(group=group_val)
                branches = branch_query.all()
                branch_ids = [b.id for b in branches]

                if branches:
                    if "Báo cáo Tuần" in report_type:
                        scores = db_session.query(WeeklyScore).filter(WeeklyScore.week == time_val, WeeklyScore.branch_id.in_(branch_ids)).all()
                        if scores:
                            start_date_str = scores[0].start_date or "…"
                            end_date_str = scores[0].end_date or "…"
                        
                        for branch in branches:
                            sc = next((s for s in scores if s.branch_id == branch.id), None)
                            gvcn_val = branch.gvcn if branch.gvcn else ""
                            so_diem_tot = calculate_trimmed_good_points_web(db_session, time_val, branch.name, branch.group, max_mon, max_tot) if sc else 0
                            data_list.append({
                                "Chi đoàn": branch.name,
                                "Nhóm": branch.group or "Nhóm 1",
                                "Sĩ số": branch.si_so,
                                "Xếp loại": getattr(sc, 'week_rating', '-') if sc else '-',
                                "Điểm Trừ VP": getattr(sc, 'score_tru', 0) if sc else 0,
                                "Số Điểm Tốt": so_diem_tot,
                                "Tổng Điểm": getattr(sc, 'total_score', 0) if sc else 0,
                                "Ghi chú VP": getattr(sc, 'note', '') if sc else 'Chưa nhập điểm',
                                "Giáo viên chủ nhiệm": gvcn_val
                            })

                    elif "Báo cáo Tháng" in report_type:
                        month_scores = db_session.query(MonthlyRecord).filter(MonthlyRecord.month_name == time_val, MonthlyRecord.school_year_id == selected_year.id, MonthlyRecord.branch_id.in_(branch_ids)).all()
                        weekly_scores_all = db_session.query(WeeklyScore).filter(WeeklyScore.branch_id.in_(branch_ids)).all()

                        for branch in branches:
                            m_sc = next((m for m in month_scores if m.branch_id == branch.id), None)
                            gvcn_val = branch.gvcn if branch.gvcn else ""
                            if m_sc:
                                ghi_chu_gop = []
                                tong_diem_tru = 0
                                tong_diem_tot = 0
                                if getattr(m_sc, 'weeks_used', None):
                                    weeks = [w.strip() for w in m_sc.weeks_used.split(",")]
                                    branch_weeks = [w for w in weekly_scores_all if w.branch_id == branch.id and w.week in weeks]
                                    for bw in branch_weeks:
                                        tong_diem_tru += (bw.score_tru or 0)
                                        if bw.note and bw.note.strip():
                                            ghi_chu_gop.append(f"[{bw.week.replace('Tuần ', 'T')}] {bw.note.strip()}")
                                    for w in weeks:
                                        tong_diem_tot += calculate_trimmed_good_points_web(db_session, w, branch.name, branch.group, max_mon, max_tot)
                                
                                xep_loai = getattr(m_sc, 'rating', 'Tốt')
                                data_list.append({
                                    "Chi đoàn": branch.name, "Nhóm": branch.group or "Nhóm 1", "Sĩ số": branch.si_so,
                                    "Xếp loại": xep_loai, "Điểm Trừ VP": tong_diem_tru, "Số Điểm Tốt": tong_diem_tot,
                                    "Tổng Điểm": m_sc.total_score, "Ghi chú VP": " | ".join(ghi_chu_gop), "Giáo viên chủ nhiệm": gvcn_val
                                })
                            else:
                                data_list.append({
                                    "Chi đoàn": branch.name, "Nhóm": branch.group or "Nhóm 1", "Sĩ số": branch.si_so,
                                    "Xếp loại": "-", "Điểm Trừ VP": 0, "Số Điểm Tốt": 0, "Tổng Điểm": 0,
                                    "Ghi chú VP": "Chưa tổng hợp tháng", "Giáo viên chủ nhiệm": gvcn_val
                                })

                    elif "Báo cáo Học kỳ" in report_type:
                        if time_val == "Học kỳ 1": target_weeks = [f"Tuần {i}" for i in range(1, 19)]
                        elif time_val == "Học kỳ 2": target_weeks = [f"Tuần {i}" for i in range(19, 38)]
                        elif time_val == "Cả năm học": target_weeks = [f"Tuần {i}" for i in range(1, 38)]
                        else: target_weeks = []

                        scores = db_session.query(WeeklyScore).filter(WeeklyScore.week.in_(target_weeks), WeeklyScore.branch_id.in_(branch_ids)).all() if target_weeks else []
                        for branch in branches:
                            branch_scores = [sc for sc in scores if sc.branch_id == branch.id]
                            gvcn_val = branch.gvcn if branch.gvcn else ""
                            if branch_scores:
                                tong_diem = sum((sc.total_score or 0) for sc in branch_scores)
                                tong_diem_tru_vp = sum((sc.score_tru or 0) for sc in branch_scores)
                                ghi_chu_gop = [f"[{sc.week.replace('Tuần ', 'T')}] {sc.note.strip()}" for sc in branch_scores if sc.note and sc.note.strip()]
                                tong_diem_tot = sum(calculate_trimmed_good_points_web(db_session, w, branch.name, branch.group, max_mon, max_tot) for w in target_weeks)
                                
                                data_list.append({
                                    "Chi đoàn": branch.name, "Nhóm": branch.group or "Nhóm 1", "Sĩ số": branch.si_so,
                                    "Xếp loại": "Tốt" if tong_diem >= 90 else "Khá", "Điểm Trừ VP": tong_diem_tru_vp,
                                    "Số Điểm Tốt": tong_diem_tot, "Tổng Điểm": round(tong_diem, 2),
                                    "Ghi chú VP": " | ".join(ghi_chu_gop), "Giáo viên chủ nhiệm": gvcn_val
                                })
                            else:
                                data_list.append({
                                    "Chi đoàn": branch.name, "Nhóm": branch.group or "Nhóm 1", "Sĩ số": branch.si_so,
                                    "Xếp loại": "-", "Điểm Trừ VP": 0, "Số Điểm Tốt": 0, "Tổng Điểm": 0,
                                    "Ghi chú VP": "Chưa có dữ liệu", "Giáo viên chủ nhiệm": gvcn_val
                                })

            if data_list:
                df = pd.DataFrame(data_list)
                if "Nhóm" in df.columns:
                    df["Hạng"] = df.groupby("Nhóm")["Tổng Điểm"].rank(method="min", ascending=False).astype(int)
                    df = df.sort_values(by=["Nhóm", "Hạng"])
                data_list = df.to_dict('records')

            return render_template(
                'templates_report.html',
                years=years,
                selected_year_name=selected_year_name,
                report_type=report_type,
                time_val=time_val,
                group_val=group_val,
                time_options=time_options,
                data_list=data_list,
                start_date=start_date_str,
                end_date=end_date_str
            )
    except Exception as e:
        import traceback; traceback.print_exc()
        flash(f"Lỗi tải biểu mẫu báo cáo: {e}", "error")
        return redirect(url_for('dashboard'))
    
# ==========================================
# MODULE: XUẤT EXCEL BIỂU MẪU (A4 DỌC, CHIA NHÓM)
# ==========================================
@app.route('/preview_templates_report', methods=['GET', 'POST'])
def preview_templates_report():
    try:
        # Hỗ trợ nhận dữ liệu linh hoạt từ cả GET và POST
        if request.method == 'POST':
            report_type = request.form.get('report_type', 'Báo cáo Tuần')
            time_val = request.form.get('time_val', 'Tuần 1')
            year_name = request.form.get('year_name', '')
            group_val = request.form.get('group_val', 'Tất cả')
        else:
            report_type = request.args.get('report_type', 'Báo cáo Tuần')
            time_val = request.args.get('time_val', 'Tuần 1')
            year_name = request.args.get('year_name', '')
            group_val = request.args.get('group_val', 'Tất cả')

        with session_scope() as db_session:
            selected_year = db_session.query(SchoolYear).filter_by(name=year_name).first()
            if not selected_year:
                flash("Không tìm thấy năm học!", "error")
                return redirect(url_for('templates_report'))

            max_tot, max_mon = 14, 4
            try:
                settings = db_session.query(ScoreSettings).filter_by(school_year_id=selected_year.id).first()
                if settings:
                    if hasattr(settings, 'max_diem_tot'): max_tot = int(settings.max_diem_tot)
                    if hasattr(settings, 'max_diem_mon'): max_mon = int(settings.max_diem_mon)
            except Exception: pass

            branch_query = db_session.query(Branch).filter_by(school_year_id=selected_year.id)
            if group_val != "Tất cả":
                if group_val.startswith("Khối"):
                    khoi = group_val.replace("Khối", "").strip()
                    from sqlalchemy import or_
                    branch_query = branch_query.filter(or_(Branch.name.startswith(khoi), Branch.name.like(f"% {khoi}%")))
                else:
                    branch_query = branch_query.filter_by(group=group_val)
            branches = branch_query.all()
            branch_ids = [b.id for b in branches]

            data_list = []
            start_str, end_str = "…", "…"

            if "Báo cáo Tuần" in report_type:
                scores = db_session.query(WeeklyScore).filter(WeeklyScore.week == time_val, WeeklyScore.branch_id.in_(branch_ids)).all()
                if scores:
                    start_str = scores[0].start_date or "…"
                    end_str = scores[0].end_date or "…"
                for branch in branches:
                    sc = next((s for s in scores if s.branch_id == branch.id), None)
                    gvcn_val = branch.gvcn if branch.gvcn else ""
                    so_diem_tot = calculate_trimmed_good_points_web(db_session, time_val, branch.name, branch.group, max_mon, max_tot) if sc else 0
                    data_list.append({
                        "Chi đoàn": branch.name, "Nhóm": branch.group or "Nhóm 1", "Sĩ số": branch.si_so,
                        "Điểm Trừ VP": int(float(getattr(sc, 'score_tru', 0) or 0)), 
                        "Số Điểm Tốt": int(float(so_diem_tot or 0)),
                        "Tổng Điểm": int(float(getattr(sc, 'total_score', 0) or 0)),
                        "Giáo viên chủ nhiệm": gvcn_val
                    })
            elif "Báo cáo Tháng" in report_type:
                month_scores = db_session.query(MonthlyRecord).filter(MonthlyRecord.month_name == time_val, MonthlyRecord.school_year_id == selected_year.id, MonthlyRecord.branch_id.in_(branch_ids)).all()
                weekly_scores_all = db_session.query(WeeklyScore).filter(WeeklyScore.branch_id.in_(branch_ids)).all()
                for branch in branches:
                    m_sc = next((m for m in month_scores if m.branch_id == branch.id), None)
                    gvcn_val = branch.gvcn if branch.gvcn else ""
                    if m_sc:
                        tong_diem_tru = 0; tong_diem_tot = 0
                        if getattr(m_sc, 'weeks_used', None):
                            weeks = [w.strip() for w in m_sc.weeks_used.split(",")]
                            branch_weeks = [w for w in weekly_scores_all if w.branch_id == branch.id and w.week in weeks]
                            for bw in branch_weeks:
                                tong_diem_tru += (bw.score_tru or 0)
                            for w in weeks:
                                tong_diem_tot += calculate_trimmed_good_points_web(db_session, w, branch.name, branch.group, max_mon, max_tot)
                        data_list.append({
                            "Chi đoàn": branch.name, "Nhóm": branch.group or "Nhóm 1", "Sĩ số": branch.si_so,
                            "Điểm Trừ VP": int(float(tong_diem_tru or 0)), 
                            "Số Điểm Tốt": int(float(tong_diem_tot or 0)), 
                            "Tổng Điểm": int(float(m_sc.total_score or 0)),
                            "Giáo viên chủ nhiệm": gvcn_val
                        })
                    else:
                        data_list.append({
                            "Chi đoàn": branch.name, "Nhóm": branch.group or "Nhóm 1", "Sĩ số": branch.si_so,
                            "Điểm Trừ VP": 0, "Số Điểm Tốt": 0, "Tổng Điểm": 0,
                            "Giáo viên chủ nhiệm": gvcn_val
                        })
            elif "Báo cáo Học kỳ" in report_type:
                if time_val == "Học kỳ 1": target_weeks = [f"Tuần {i}" for i in range(1, 19)]
                elif time_val == "Học kỳ 2": target_weeks = [f"Tuần {i}" for i in range(19, 38)]
                elif time_val == "Cả năm học": target_weeks = [f"Tuần {i}" for i in range(1, 38)]
                else: target_weeks = []

                scores = db_session.query(WeeklyScore).filter(WeeklyScore.week.in_(target_weeks), WeeklyScore.branch_id.in_(branch_ids)).all() if target_weeks else []
                for branch in branches:
                    branch_scores = [sc for sc in scores if sc.branch_id == branch.id]
                    gvcn_val = branch.gvcn if branch.gvcn else ""
                    if branch_scores:
                        tong_diem = sum((sc.total_score or 0) for sc in branch_scores)
                        tong_diem_tru_vp = sum((sc.score_tru or 0) for sc in branch_scores)
                        tong_diem_tot = sum(calculate_trimmed_good_points_web(db_session, w, branch.name, branch.group, max_mon, max_tot) for w in target_weeks)
                        data_list.append({
                            "Chi đoàn": branch.name, "Nhóm": branch.group or "Nhóm 1", "Sĩ số": branch.si_so,
                            "Điểm Trừ VP": int(float(tong_diem_tru_vp or 0)), 
                            "Số Điểm Tốt": int(float(tong_diem_tot or 0)), 
                            "Tổng Điểm": int(float(tong_diem or 0)),
                            "Giáo viên chủ nhiệm": gvcn_val
                        })
                    else:
                        data_list.append({
                            "Chi đoàn": branch.name, "Nhóm": branch.group or "Nhóm 1", "Sĩ số": branch.si_so,
                            "Điểm Trừ VP": 0, "Số Điểm Tốt": 0, "Tổng Điểm": 0,
                            "Giáo viên chủ nhiệm": gvcn_val
                        })

            df = pd.DataFrame(data_list)
            if not df.empty and "Nhóm" in df.columns:
                df["Hạng"] = df.groupby("Nhóm")["Tổng Điểm"].rank(method="min", ascending=False).astype(int)
                df = df.sort_values(by=["Nhóm", "Hạng"])
                data_list = df.to_dict('records')

            return render_template(
                'preview_templates_excel.html',
                report_type=report_type,
                time_val=time_val,
                year_name=year_name,
                group_val=group_val,
                start_str=start_str,
                end_str=end_str,
                data_list=data_list
            )
    except Exception as e:
        import traceback; traceback.print_exc()
        flash(f"Lỗi xem trước: {e}", "error")
        return redirect(url_for('templates_report'))
# ==========================================
# MODULE: XUẤT EXCEL BIỂU MẪU CHÍNH THỨC (A4 DỌC, CHIA NHÓM)
# ==========================================
@app.route('/export_templates_excel', methods=['POST'])
def export_templates_excel():
    try:
        report_type = request.form.get('report_type', 'Báo cáo Tuần')
        time_val = request.form.get('time_val', 'Tuần 1')
        year_name = request.form.get('year_name', '')
        group_val = request.form.get('group_val', 'Tất cả')

        with session_scope() as db_session:
            selected_year = db_session.query(SchoolYear).filter_by(name=year_name).first()
            if not selected_year:
                flash("Không tìm thấy năm học!", "error")
                return redirect(url_for('templates_report'))

            max_tot, max_mon = 14, 4
            try:
                settings = db_session.query(ScoreSettings).filter_by(school_year_id=selected_year.id).first()
                if settings:
                    if hasattr(settings, 'max_diem_tot'): max_tot = int(settings.max_diem_tot)
                    if hasattr(settings, 'max_diem_mon'): max_mon = int(settings.max_diem_mon)
            except Exception: pass

            branch_query = db_session.query(Branch).filter_by(school_year_id=selected_year.id)
            if group_val != "Tất cả":
                if group_val.startswith("Khối"):
                    khoi = group_val.replace("Khối", "").strip()
                    from sqlalchemy import or_
                    branch_query = branch_query.filter(or_(Branch.name.startswith(khoi), Branch.name.like(f"% {khoi}%")))
                else:
                    branch_query = branch_query.filter_by(group=group_val)
            branches = branch_query.all()
            branch_ids = [b.id for b in branches]

            if not branches:
                flash("Không có dữ liệu chi đoàn để xuất!", "error")
                return redirect(url_for('templates_report'))

            data_list = []
            start_str, end_str = "…", "…"

            if "Báo cáo Tuần" in report_type:
                scores = db_session.query(WeeklyScore).filter(WeeklyScore.week == time_val, WeeklyScore.branch_id.in_(branch_ids)).all()
                if scores:
                    start_str = scores[0].start_date or "…"
                    end_str = scores[0].end_date or "…"
                for branch in branches:
                    sc = next((s for s in scores if s.branch_id == branch.id), None)
                    gvcn_val = branch.gvcn if branch.gvcn else ""
                    so_diem_tot = calculate_trimmed_good_points_web(db_session, time_val, branch.name, branch.group, max_mon, max_tot) if sc else 0
                    data_list.append({
                        "Chi đoàn": branch.name, "Nhóm": branch.group or "Nhóm 1", "Sĩ số": branch.si_so,
                        "Điểm Trừ VP": int(float(getattr(sc, 'score_tru', 0) or 0)), 
                        "Số Điểm Tốt": int(float(so_diem_tot or 0)),
                        "Tổng Điểm": int(float(getattr(sc, 'total_score', 0) or 0)),
                        "Ghi chú VP": getattr(sc, 'note', '') if sc else 'Chưa nhập điểm',
                        "Giáo viên chủ nhiệm": gvcn_val
                    })
            elif "Báo cáo Tháng" in report_type:
                month_scores = db_session.query(MonthlyRecord).filter(MonthlyRecord.month_name == time_val, MonthlyRecord.school_year_id == selected_year.id, MonthlyRecord.branch_id.in_(branch_ids)).all()
                weekly_scores_all = db_session.query(WeeklyScore).filter(WeeklyScore.branch_id.in_(branch_ids)).all()
                for branch in branches:
                    m_sc = next((m for m in month_scores if m.branch_id == branch.id), None)
                    gvcn_val = branch.gvcn if branch.gvcn else ""
                    if m_sc:
                        ghi_chu_gop = []; tong_diem_tru = 0; tong_diem_tot = 0
                        if getattr(m_sc, 'weeks_used', None):
                            weeks = [w.strip() for w in m_sc.weeks_used.split(",")]
                            branch_weeks = [w for w in weekly_scores_all if w.branch_id == branch.id and w.week in weeks]
                            for bw in branch_weeks:
                                tong_diem_tru += (bw.score_tru or 0)
                                if bw.note and bw.note.strip():
                                    ghi_chu_gop.append(f"[{bw.week.replace('Tuần ', 'T')}] {bw.note.strip()}")
                            for w in weeks:
                                tong_diem_tot += calculate_trimmed_good_points_web(db_session, w, branch.name, branch.group, max_mon, max_tot)
                        data_list.append({
                            "Chi đoàn": branch.name, "Nhóm": branch.group or "Nhóm 1", "Sĩ số": branch.si_so,
                            "Điểm Trừ VP": int(float(tong_diem_tru or 0)), 
                            "Số Điểm Tốt": int(float(tong_diem_tot or 0)), 
                            "Tổng Điểm": int(float(m_sc.total_score or 0)),
                            "Ghi chú VP": " | ".join(ghi_chu_gop), "Giáo viên chủ nhiệm": gvcn_val
                        })
                    else:
                        data_list.append({
                            "Chi đoàn": branch.name, "Nhóm": branch.group or "Nhóm 1", "Sĩ số": branch.si_so,
                            "Điểm Trừ VP": 0, "Số Điểm Tốt": 0, "Tổng Điểm": 0,
                            "Ghi chú VP": "Chưa tổng hợp tháng", "Giáo viên chủ nhiệm": gvcn_val
                        })
            elif "Báo cáo Học kỳ" in report_type:
                if time_val == "Học kỳ 1": target_weeks = [f"Tuần {i}" for i in range(1, 19)]
                elif time_val == "Học kỳ 2": target_weeks = [f"Tuần {i}" for i in range(19, 38)]
                elif time_val == "Cả năm học": target_weeks = [f"Tuần {i}" for i in range(1, 38)]
                else: target_weeks = []

                scores = db_session.query(WeeklyScore).filter(WeeklyScore.week.in_(target_weeks), WeeklyScore.branch_id.in_(branch_ids)).all() if target_weeks else []
                for branch in branches:
                    branch_scores = [sc for sc in scores if sc.branch_id == branch.id]
                    gvcn_val = branch.gvcn if branch.gvcn else ""
                    if branch_scores:
                        tong_diem = sum((sc.total_score or 0) for sc in branch_scores)
                        tong_diem_tru_vp = sum((sc.score_tru or 0) for sc in branch_scores)
                        ghi_chu_gop = [f"[{sc.week.replace('Tuần ', 'T')}] {sc.note.strip()}" for sc in branch_scores if sc.note and sc.note.strip()]
                        tong_diem_tot = sum(calculate_trimmed_good_points_web(db_session, w, branch.name, branch.group, max_mon, max_tot) for w in target_weeks)
                        data_list.append({
                            "Chi đoàn": branch.name, "Nhóm": branch.group or "Nhóm 1", "Sĩ số": branch.si_so,
                            "Điểm Trừ VP": int(float(tong_diem_tru_vp or 0)), 
                            "Số Điểm Tốt": int(float(tong_diem_tot or 0)), 
                            "Tổng Điểm": int(float(tong_diem or 0)),
                            "Ghi chú VP": " | ".join(ghi_chu_gop), "Giáo viên chủ nhiệm": gvcn_val
                        })
                    else:
                        data_list.append({
                            "Chi đoàn": branch.name, "Nhóm": branch.group or "Nhóm 1", "Sĩ số": branch.si_so,
                            "Điểm Trừ VP": 0, "Số Điểm Tốt": 0, "Tổng Điểm": 0,
                            "Ghi chú VP": "Chưa có dữ liệu", "Giáo viên chủ nhiệm": gvcn_val
                        })

            df = pd.DataFrame(data_list)
            if not df.empty and "Nhóm" in df.columns:
                df["Hạng"] = df.groupby("Nhóm")["Tổng Điểm"].rank(method="min", ascending=False).astype(int)
                df = df.sort_values(by=["Nhóm", "Hạng"])

            # Khởi tạo file Excel với định dạng A4 Dọc
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Bao_Cao_Thi_Dua"

            # Cấu hình khổ giấy A4 Dọc và chế độ xem trước
            ws.page_setup.paperSize = ws.PAPERSIZE_A4
            ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT  # <--- Định dạng A4 Dọc
            ws.sheet_properties.pageSetUpPr.fitToPage = True
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 0
            ws.views.sheetView[0].view = "pageBreakPreview"

            # Thiết lập Header văn bản hành chính
            ws['B1'] = "BCH ĐOÀN XÃ THIỆN HƯNG"
            ws['B1'].font = Font(name='Times New Roman', bold=True, size=11)
            ws['B1'].alignment = Alignment(horizontal='center')

            ws['G1'] = "ĐOÀN TNCS HỒ CHÍ MINH"
            ws['G1'].font = Font(name='Times New Roman', bold=True, size=11)
            ws['G1'].alignment = Alignment(horizontal='center')

            ws.merge_cells('B2:C2')
            ws['B2'] = "ĐOÀN TRƯỜNG THPT THANH HÒA"
            ws['B2'].font = Font(name='Times New Roman', bold=True, size=11)
            ws['B2'].alignment = Alignment(horizontal='center')

            ws.merge_cells('F2:H2')
            ws['F2'] = "-----***-----"
            ws['F2'].font = Font(name='Times New Roman', bold=True, size=11)
            ws['F2'].alignment = Alignment(horizontal='center')

            ws.merge_cells('B4:H4')
            ws['B4'] = f"BẢNG {report_type.upper()} - {time_val.upper()}"
            ws['B4'].font = Font(name='Times New Roman', bold=True, size=15)
            ws['B4'].alignment = Alignment(horizontal='center', vertical='center')

            if not start_str or start_str == "…": start_str = "........"
            if not end_str or end_str == "…": end_str = "........"

            ws.merge_cells('B5:H5')
            ws['B5'] = f"(Thời gian: Từ ngày {start_str} đến ngày {end_str})"
            ws['B5'].font = Font(name='Times New Roman', italic=True, size=11)
            ws['B5'].alignment = Alignment(horizontal='center', vertical='center')

            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            header_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
            group_fill = PatternFill(start_color="D1E7DD", end_color="D1E7DD", fill_type="solid")

            current_row = 7
            # Đã lược bỏ 'Ghi chú VP'
            columns_to_print = ['STT', 'Chi đoàn', 'Sĩ số', 'Điểm Trừ VP', 'Số Điểm Tốt', 'Hạng', 'Tổng Điểm', 'Giáo viên chủ nhiệm']
            
            grouped = df.groupby('Nhóm')
            for name, group in grouped:
                ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=8)
                g_cell = ws.cell(row=current_row, column=1, value=f"PHÂN BẢNG THI ĐUA: {str(name).upper()}")
                g_cell.font = Font(name='Times New Roman', bold=True, size=12, color="0F5132")
                g_cell.fill = group_fill
                g_cell.alignment = Alignment(horizontal='left', vertical='center')
                for c in range(1, 9):
                    ws.cell(row=current_row, column=c).border = thin_border
                current_row += 1

                for col_num, header_name in enumerate(columns_to_print, 1):
                    cell = ws.cell(row=current_row, column=col_num, value=header_name)
                    cell.font = Font(name='Times New Roman', bold=True, size=11)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.border = thin_border
                    cell.fill = header_fill
                current_row += 1

                for stt, (_, row_data) in enumerate(group.iterrows(), 1):
                    cell_stt = ws.cell(row=current_row, column=1, value=stt)
                    cell_stt.border = thin_border
                    cell_stt.alignment = Alignment(horizontal='center')

                    # Đã lược bỏ 'Ghi chú VP' khỏi danh sách khóa dữ liệu
                    col_keys = ['Chi đoàn', 'Sĩ số', 'Điểm Trừ VP', 'Số Điểm Tốt', 'Hạng', 'Tổng Điểm', 'Giáo viên chủ nhiệm']
                    for col_num, col_name in enumerate(col_keys, 2):
                        cell = ws.cell(row=current_row, column=col_num)
                        val = row_data.get(col_name, "")
                        if isinstance(val, float) and val.is_integer():
                            val = int(val)
                        cell.value = val
                        cell.font = Font(name='Times New Roman', size=11)
                        cell.border = thin_border
                        
                        if col_name in ['Sĩ số', 'Hạng', 'Tổng Điểm', 'Số Điểm Tốt', 'Điểm Trừ VP']:
                            cell.alignment = Alignment(horizontal='center')
                        else:
                            cell.alignment = Alignment(horizontal='left')
                    current_row += 1
                current_row += 1

            # Điều chỉnh lại độ rộng các cột tối ưu chuẩn A4 Dọc khi không có cột ghi chú
            widths = {'A': 6, 'B': 15, 'C': 9, 'D': 14, 'E': 14, 'F': 10, 'G': 12, 'H': 28}
            for col_letter, width in widths.items():
                ws.column_dimensions[col_letter].width = width

            log_system_action("XUẤT EXCEL", f"Xuất biểu mẫu A4 dọc chia nhóm: {report_type} - {time_val}")
            out = io.BytesIO()
            wb.save(out)
            out.seek(0)

            file_name = f"Bao_Cao_{report_type}_{time_val}.xlsx".replace(" ", "_")
            return send_file(out, download_name=file_name, as_attachment=True)
            
    except Exception as e:
        import traceback; traceback.print_exc()
        flash(f"Lỗi xuất Excel: {e}", "error")
        return redirect(url_for('templates_report'))
    

# ==========================================
# MODULE: SAO LƯU, PHỤC HỒI & QUẢN LÝ DỮ LIỆU (CHỈ DÀNH CHO ADMIN)
# ==========================================
def perform_backup_internal(actor="Hệ thống"):
    import glob, shutil
    # [BẢN VÁ LỖI MÚI GIỜ]: Khai báo múi giờ Việt Nam
    from datetime import datetime, timezone, timedelta
    vn_tz = timezone(timedelta(hours=7))
    
    backup_dir = "backups"
    os.makedirs(backup_dir, exist_ok=True)
    
    # Ép lấy giờ Việt Nam để đặt tên file backup cho chuẩn xác
    timestamp = datetime.now(vn_tz).strftime('%Y%m%d_%H%M%S')
    backup_filename = f"Data_ThiDua_Backup_{timestamp}"
    backup_path = os.path.join(backup_dir, backup_filename)
    temp_dir = os.path.join(backup_dir, f"temp_{timestamp}")
    os.makedirs(temp_dir, exist_ok=True)
    
    # Gom dữ liệu
    for folder in ["data", "config", "database"]:
        if os.path.exists(folder):
            shutil.copytree(folder, os.path.join(temp_dir, folder))
            
    # Nén zip
    shutil.make_archive(backup_path, 'zip', temp_dir)
    shutil.rmtree(temp_dir) # Dọn dẹp thư mục tạm
    
    # THUẬT TOÁN 1: Dọn rác tự động - Quét và chỉ giữ lại đúng 20 bản sao lưu mới nhất
    files = glob.glob(os.path.join(backup_dir, "*.zip"))
    files.sort(key=os.path.getmtime, reverse=True)
    if len(files) > 20:
        for old_file in files[20:]:
            try: os.remove(old_file)
            except: pass
            
    # Ép lấy giờ Việt Nam để in ra log hệ thống
    print(f"[{datetime.now(vn_tz).strftime('%H:%M:%S')}] {actor} đã tạo bản sao lưu: {backup_filename}.zip")
    return f"{backup_filename}.zip"

# THUẬT TOÁN 3: Lập lịch sao lưu ngầm (Chạy nền)
def background_auto_backup():
    import time
    # [BẢN VÁ LỖI MÚI GIỜ]: Khai báo thư viện và múi giờ Việt Nam
    from datetime import datetime, timezone, timedelta
    vn_tz = timezone(timedelta(hours=7))
    
    while True:
        # Ép Robot lấy đồng hồ theo giờ Việt Nam
        now_vn = datetime.now(vn_tz) 
        
        # Nếu là Chủ nhật (weekday == 6) và thời gian rơi vào 23h55' đêm (Giờ VN)
        if now_vn.weekday() == 6 and now_vn.hour == 23 and now_vn.minute == 55:
            try:
                perform_backup_internal(actor="BOT Lịch ngầm Tự động")
            except Exception as e:
                print("Lỗi auto backup:", e)
            time.sleep(3600) # Ngủ đông 1 tiếng để tránh chạy đúp lệnh trong cùng 1 đêm
        time.sleep(45) # Quét đồng hồ mỗi 45 giây

@app.route('/backup')
def backup_manager():
    if session.get('role') not in ['Quản trị viên', 'Admin', 'Bí thư Đoàn trường', 'Bí thư']:
        flash("Bạn không có quyền truy cập chức năng này!", "error")
        return redirect(url_for('dashboard'))
        
    import glob
    backup_dir = "backups"
    os.makedirs(backup_dir, exist_ok=True)
    
    files = glob.glob(os.path.join(backup_dir, "*.zip"))
    backups = []
    for f in files:
        stat = os.stat(f)
        backups.append({
            'filename': os.path.basename(f),
            'size': round(stat.st_size / 1024 / 1024, 2),
            'time': datetime.fromtimestamp(stat.st_mtime).strftime('%d/%m/%Y - %H:%M:%S'),
            'raw_time': stat.st_mtime
        })
        
    backups.sort(key=lambda x: x['raw_time'], reverse=True)
    return render_template('backup.html', backups=backups)

@app.route('/create_backup', methods=['POST'])
def create_backup():
    if session.get('role') not in ['Quản trị viên', 'Admin', 'Bí thư Đoàn trường', 'Bí thư']:
        return redirect(url_for('dashboard'))
    try:
        filename = perform_backup_internal(actor=session.get('username'))
        log_system_action("SAO LƯU DỮ LIỆU", f"Đã tạo bản sao lưu hệ thống: {filename}")
        flash("✅ Đã tạo sao lưu, đóng gói và tự động dọn rác thành công!", "success")
    except Exception as e:
        flash(f"Lỗi tạo sao lưu: {e}", "error")
    return redirect(url_for('backup_manager'))

# THUẬT TOÁN 2: Phục hồi 1 chạm (1-Click Restore)
@app.route('/restore_backup/<filename>', methods=['POST'])
def restore_backup(filename):
    if session.get('role') not in ['Quản trị viên', 'Admin', 'Bí thư Đoàn trường', 'Bí thư']:
        return redirect(url_for('dashboard'))
    try:
        import shutil
        import zipfile
        import os
        import time
        
        file_path = os.path.join("backups", filename)
        
        if not os.path.exists(file_path):
            flash("Không tìm thấy file sao lưu trên máy chủ!", "error")
            return redirect(url_for('backup_manager'))
            
        temp_extract = os.path.join("backups", "temp_restore")
        os.makedirs(temp_extract, exist_ok=True)
        
        # 1. Xả nén file Zip
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(temp_extract)
            
        # --- [BẢN VÁ LỖI WINERROR 32]: NGẮT KẾT NỐI CSDL TRƯỚC KHI XÓA ---
        try:
            from database.database import engine
            engine.dispose()  # Ép SQLAlchemy nhả file thi_dua.db ra khỏi RAM
            time.sleep(1.5)   # Đợi 1.5 giây để HĐH Windows thu hồi file handle
        except Exception as e:
            print("Cảnh báo ngắt kết nối DB:", e)
        # -----------------------------------------------------------------
            
        # 2. Xóa dữ liệu cũ và ghi đè dữ liệu mới
        for folder in ["data", "config", "database"]:
            src = os.path.join(temp_extract, folder)
            if os.path.exists(src):
                if os.path.exists(folder):
                    shutil.rmtree(folder, ignore_errors=True) # Xóa thư mục hiện tại
                    time.sleep(0.5) # Nghỉ nhịp để HĐH làm mới cây thư mục
                    
                    # Cố gắng copy thư mục, nếu kẹt file thì copy đè từng file
                    if not os.path.exists(folder):
                        shutil.copytree(src, folder)
                    else:
                        for root, dirs, files in os.walk(src):
                            for f in files:
                                src_file = os.path.join(root, f)
                                dst_file = os.path.join(folder, os.path.relpath(src_file, src))
                                os.makedirs(os.path.dirname(dst_file), exist_ok=True)
                                shutil.copy2(src_file, dst_file)
                else:
                    shutil.copytree(src, folder)
                    
        shutil.rmtree(temp_extract, ignore_errors=True) # Dọn rác xả nén
        
        log_system_action("PHỤC HỒI DỮ LIỆU", f"Đã phục hồi hệ thống từ bản sao lưu: {filename}")
        flash(f"🔄 ĐÃ PHỤC HỒI THÀNH CÔNG DỮ LIỆU TỪ FILE {filename}!", "success")
    except Exception as e:
        import traceback
        traceback.print_exc()
        flash(f"Lỗi phục hồi dữ liệu: {e}", "error")
        
    return redirect(url_for('backup_manager'))

@app.route('/download_backup/<filename>')
def download_backup(filename):
    if session.get('role') not in ['Quản trị viên', 'Admin', 'Bí thư Đoàn trường', 'Bí thư']:
        return redirect(url_for('dashboard'))
        
    file_path = os.path.join("backups", filename)
    if os.path.exists(file_path):
        log_system_action("SAO LƯU DỮ LIỆU", f"Tải xuống file sao lưu: {filename}")
        return send_file(file_path, as_attachment=True)
    else:
        flash("Không tìm thấy file sao lưu trên máy chủ!", "error")
        return redirect(url_for('backup_manager'))

@app.route('/delete_backup/<filename>', methods=['POST'])
def delete_backup(filename):
    if session.get('role') not in ['Quản trị viên', 'Admin', 'Bí thư Đoàn trường', 'Bí thư']:
        return redirect(url_for('dashboard'))
        
    file_path = os.path.join("backups", filename)
    if os.path.exists(file_path):
        os.remove(file_path)
        log_system_action("SAO LƯU DỮ LIỆU", f"Đã xóa bản sao lưu: {filename}")
        flash(f"Đã xóa file {filename} thành công!", "success")
    else:
        flash("Không tìm thấy file trên máy chủ!", "error")
    return redirect(url_for('backup_manager'))

# ==========================================
# API TỰ ĐỘNG CẤP TÀI KHOẢN HÀNG LOẠT CHO GVCN
# ==========================================
@app.route('/auto_generate_gvcn', methods=['POST'])
def auto_generate_gvcn():
    if session.get('role') not in ['Quản trị viên', 'Admin', 'Bí thư Đoàn trường', 'Bí thư']:
        flash("Chỉ Admin mới có quyền thực hiện!", "error")
        return redirect(url_for('users'))
        
    try:
        with session_scope() as db_session:
            active_year = db_session.query(SchoolYear).filter_by(is_active=True).first()
            if not active_year: return redirect(url_for('users'))
                
            branches = db_session.query(Branch).filter_by(school_year_id=active_year.id).all()
            count = 0
            
            for b in branches:
                username = b.name.strip().upper()
                exist = db_session.query(User).filter_by(username=username).first()
                
                if not exist:
                    # --- NÂNG CẤP: LẤY TÊN THẬT CỦA GVCN TỪ CHI ĐOÀN ---
                    # Nếu lớp đã có tên GVCN thì lấy tên đó, nếu trống thì mới dùng tạm "GVCN Lớp ..."
                    ten_gvcn = b.gvcn.strip() if b.gvcn and b.gvcn.strip() else f"GVCN Lớp {username}"
                    
                    new_user = User(
                        username=username, password_hash=username, # Pass mặc định = Tên lớp
                        full_name=ten_gvcn, role=UserRole.GVCN, is_active=True
                    )
                    db_session.add(new_user)
                    count += 1
                    try: sync_account_to_json(username, new_user.full_name, username, "Giáo viên chủ nhiệm", True)
                    except: pass
            
            log_system_action("CẤP TÀI KHOẢN", f"Đã tự động tạo {count} tài khoản GVCN.")
            flash(f"✅ Đã cấp phát tự động {count} tài khoản cho GVCN! (Tên đăng nhập = Mật khẩu = Tên lớp)", "success")
            return redirect(url_for('users'))
    except Exception as e:
        flash(f"Lỗi hệ thống: {e}", "error")
        return redirect(url_for('users'))
# ==========================================
# API: GVCN GỬI PHÚC KHẢO ĐIỂM (BẢN VÁ LỖI MÚI GIỜ UTC+7 TỐI THƯỢNG)
# ==========================================
@app.route('/submit_appeal', methods=['POST'])
def submit_appeal():
    if session.get('role') != 'Giáo viên chủ nhiệm': 
        return redirect(url_for('login'))
        
    score_id = request.form.get('score_id', type=int)
    reason = request.form.get('reason', '').strip()
    evidence_base64 = request.form.get('appeal_evidence_base64', '').replace(' ', '+')
    
    if not score_id or not reason: 
        return redirect(url_for('class_dashboard'))
        
    try:
        with session_scope() as db_session:
            score = db_session.query(WeeklyScore).filter_by(id=score_id).first()
            
            if not score:
                flash("Không tìm thấy dữ liệu điểm!", "error")
                return redirect(url_for('class_dashboard'))
            
            if score.is_locked:
                flash("⛔ Tuần này đã chốt cứng, không thể gửi yêu cầu phúc khảo!", "error")
                return redirect(url_for('class_dashboard'))
            
            # =========================================================================
            # [BẢN VÁ TỐI THƯỢNG]: ÉP BUỘC MÚI GIỜ VIỆT NAM (UTC+7) ĐỒNG BỘ 100%
            # =========================================================================
            from datetime import datetime, timezone, timedelta
            # Khởi tạo đối tượng múi giờ VN chuẩn
            vn_tz = timezone(timedelta(hours=7))
            # Lấy giờ hiện tại và ép chặt vào múi giờ VN
            now_vn = datetime.now(vn_tz) 
            today_vn_date = now_vn.date()
            
            # --- LUẬT KHÓA CHỦ NHẬT ---
            is_expired_dynamic = False
            if score.start_date:
                try:
                    start_d_clean = score.start_date.split()[0]
                    if "-" in start_d_clean:
                        start_date_obj = datetime.strptime(start_d_clean, '%Y-%m-%d').date()
                    else:
                        start_date_obj = datetime.strptime(start_d_clean, '%d/%m/%Y').date()
                        
                    sunday_obj = start_date_obj + timedelta(days=6)
                    if today_vn_date > sunday_obj: # Dùng ngày VN để so sánh
                        is_expired_dynamic = True
                except Exception as e:
                    pass
            
            if is_expired_dynamic:
                flash("⛔ Đã hết thời hạn! Hệ thống tự động khóa quyền khiếu nại vào ngày Chủ nhật của tuần thi đua.", "error")
                return redirect(url_for('class_dashboard'))

            # --- LUẬT KHÓA LỖI QUÁ NGÀY ---
            
            days_vn = {0: '[T2]', 1: '[T3]', 2: '[T4]', 3: '[T5]', 4: '[T6]', 5: '[T7]', 6: '[CN]'}
            # ĐÃ SỬA TẠI ĐÂY: Lấy thứ theo lịch Việt Nam thay vì giờ của Server Mỹ
            today_pfx = days_vn[now_vn.weekday()] 
            
            match_errors = re.search(r'Phúc khảo các lỗi:\s*\[(.*?)\]', reason)
            if match_errors:
                errors_str = match_errors.group(1)
                appealed_errors = [e.strip() for e in errors_str.split("] & [")]
                
                for err in appealed_errors:
                    day_match = re.search(r'\[(T[2-7](?:\s*Chiều|\s*Chieu)?|CN)\]', err, re.IGNORECASE)
                    if day_match:
                        err_day = day_match.group(0).upper()
                        # Kiểm tra xem thứ hiện tại (VD: T5) có nằm trong tên lỗi (VD: [T5 CHIỀU]) hay không
                        if today_pfx.strip('[]') not in err_day:
                            flash(f"⛔ TỪ CHỐI: Lỗi thuộc ngày {err_day} đã quá hạn! Chỉ tiếp nhận khiếu nại trong cùng ngày xảy ra vi phạm.", "error")
                            return redirect(url_for('class_dashboard'))

            # =================================================================
            # TẢI ẢNH LÊN CLOUDINARY VÀ GẮN LINK VÀO GHI CHÚ
            # =================================================================
            reason_with_img = reason 
            
            if evidence_base64:
                saved_image_url = process_and_save_evidence(evidence_base64, score.branch_id, score.week)
                if not saved_image_url:
                    flash("⛔ Có lỗi xảy ra khi tải ảnh lên đám mây Cloudinary. Vui lòng thử lại!", "error")
                    return redirect(url_for('class_dashboard'))
                
                reason_with_img = f"{reason} <br><a href='{saved_image_url}' target='_blank' style='color: #2563eb; text-decoration: none; display: inline-block; margin-top: 8px;'><i class='fa-regular fa-image'></i> <b>Xem ảnh minh chứng GVCN gửi</b></a>"

            # --- KIỂM TRA SỐ LẦN GỬI (TỐI ĐA 2 LẦN/NGÀY) THEO GIỜ VN ---
            # ĐÃ SỬA TẠI ĐÂY: Dùng now_vn thay cho datetime.now()
            today_date_str = now_vn.strftime("%d/%m/%Y") 
            now_str = now_vn.strftime("%d/%m/%Y %H:%M")  
            new_entry = f"[{now_str}] {reason_with_img}" 
            
            if score.appeal_reason:
                count_today = score.appeal_reason.count(f"[{today_date_str}")
                if count_today >= 2:
                    flash("⛔ Thầy/cô đã dùng hết 2 lượt gửi phúc khảo trong ngày hôm nay!", "error")
                    return redirect(url_for('class_dashboard'))
                
                score.appeal_reason = score.appeal_reason + " | " + new_entry
            else:
                score.appeal_reason = new_entry
            
            score.is_appealed = True
            score.appeal_response = None 
            
            log_system_action("PHÚC KHẢO", f"GVCN Lớp {score.branch.name} gửi khiếu nại.")
            flash("✅ Đã gửi Báo cáo sai sót / Phúc khảo đến Đoàn trường thành công!", "success")
            
    except Exception as e: 
        err_msg = str(e)[:150]
        flash(f"Lỗi hệ thống: {err_msg}", "error")
        import traceback; traceback.print_exc() 
        
    return redirect(url_for('class_dashboard'))

# ==========================================
# MODULE: WEB APP MOBILE DÀNH CHO SAO ĐỎ
# ==========================================

@app.route('/auto_generate_saodo', methods=['POST'])
def auto_generate_saodo():
    if session.get('role') not in ['Quản trị viên', 'Admin', 'Bí thư']: return redirect(url_for('users'))
    try:
        with session_scope() as db_session:
            stars = db_session.query(RedStar).filter_by(is_active=True).all()
            count = 0
            for s in stars:
                # Tên đăng nhập: SD + ID của Sao đỏ (VD: SD1, SD15)
                username = f"SD{s.id}"
                exist = db_session.query(User).filter_by(username=username).first()
                if not exist:
                    new_user = User(
                        username=username, password_hash="123456", # Mật khẩu mặc định: 123456
                        full_name=f"SĐ: {s.full_name}", role=UserRole.SAO_DO, is_active=True
                    )
                    db_session.add(new_user)
                    count += 1
                    try: sync_account_to_json(username, new_user.full_name, "123456", "Sao đỏ", True)
                    except: pass
            
            log_system_action("CẤP TÀI KHOẢN", f"Đã tự động tạo {count} tài khoản Sao đỏ.")
            flash(f"✅ Đã cấp phát tự động {count} tài khoản Sao đỏ! (Tài khoản: SD + ID, Mật khẩu: 123456)", "success")
            return redirect(url_for('users'))
    except Exception as e:
        flash(f"Lỗi hệ thống: {e}", "error"); return redirect(url_for('users'))
# ==========================================
# API XUẤT DANH SÁCH TÀI KHOẢN SAO ĐỎ RA EXCEL
# ==========================================
@app.route('/export_saodo_accounts')
def export_saodo_accounts():
    if session.get('role') not in ['Quản trị viên', 'Admin', 'Bí thư Đoàn trường', 'Bí thư']:
        flash("Bạn không có quyền thực hiện chức năng này!", "error")
        return redirect(url_for('users'))
        
    try:
        with session_scope() as db_session:
            # Lấy toàn bộ tài khoản Sao đỏ đang hoạt động
            sao_do_users = db_session.query(User).filter(User.role == UserRole.SAO_DO, User.is_active == True).all()
            
            if not sao_do_users:
                flash("Chưa có tài khoản Sao đỏ nào trên hệ thống!", "warning")
                return redirect(url_for('users'))
            
            import openpyxl
            from openpyxl.styles import Font, Alignment, Border, Side
            import io
            from flask import send_file
            
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "DS_Tai_Khoan_Sao_Do"
            
            ws.merge_cells('A1:E1')
            ws['A1'] = "DANH SÁCH BÀN GIAO TÀI KHOẢN APP SAO ĐỎ"
            ws['A1'].font = Font(name="Times New Roman", size=14, bold=True)
            ws['A1'].alignment = Alignment(horizontal="center")
            
            headers = ["STT", "Họ và Tên", "Chi đoàn", "Tên Đăng Nhập", "Mật Khẩu"]
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            
            for col, h in enumerate(headers, 1):
                c = ws.cell(row=3, column=col, value=h)
                c.font = Font(name="Times New Roman", size=12, bold=True)
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.border = thin_border
                
            for idx, u in enumerate(sao_do_users, 1):
                # Tách ID từ tên đăng nhập (VD: SD15 -> 15) để dò ra Tên Lớp
                branch_name = "Không rõ"
                try:
                    rs_id = int(u.username.replace('SD', ''))
                    rs = db_session.query(RedStar).filter_by(id=rs_id).first()
                    if rs and rs.branch:
                        branch_name = rs.branch.name
                except: pass
                
                # Làm sạch tên hiển thị
                real_name = u.full_name.replace("SĐ: ", "") if u.full_name else ""
                
                row_data = [idx, real_name, branch_name, u.username, u.password_hash]
                for col, val in enumerate(row_data, 1):
                    c = ws.cell(row=idx+3, column=col, value=val)
                    c.font = Font(name="Times New Roman", size=12)
                    c.border = thin_border
                    if col in [1, 3, 4, 5]: c.alignment = Alignment(horizontal="center")
                    
            ws.column_dimensions['A'].width = 8
            ws.column_dimensions['B'].width = 30
            ws.column_dimensions['C'].width = 15
            ws.column_dimensions['D'].width = 20
            ws.column_dimensions['E'].width = 15
            
            log_system_action("XUẤT EXCEL", "Xuất danh sách bàn giao tài khoản Sao đỏ")
            out = io.BytesIO()
            wb.save(out)
            out.seek(0)
            return send_file(out, download_name="Danh_Sach_Tai_Khoan_Sao_Do.xlsx", as_attachment=True)
            
    except Exception as e:
        flash(f"Lỗi xuất Excel: {e}", "error")
        return redirect(url_for('users'))
# ==========================================
# TRẠM PHÁT SÓNG THÔNG BÁO CHO APP SAO ĐỎ
# ==========================================
@app.route('/update_announcement', methods=['POST'])
def update_announcement():
    if session.get('role') not in ['Quản trị viên', 'Admin', 'Bí thư Đoàn trường', 'Bí thư']:
        flash("Bạn không có quyền phát thông báo!", "error")
        return redirect(request.referrer or url_for('dashboard'))
        
    new_text = request.form.get('announcement_text', '').strip()
    if new_text:
        import os, json
        config_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config")
        os.makedirs(config_dir, exist_ok=True)
        config_path = os.path.join(config_dir, "announcement.json")
        
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump({"text": new_text}, f, ensure_ascii=False)
            
        log_system_action("PHÁT THÔNG BÁO", f"Nội dung: {new_text[:30]}...")
        flash("✅ Đã phát lệnh điều hành bằng chữ chạy tới toàn bộ App Sao Đỏ!", "success")
        
    return redirect(request.referrer or url_for('dashboard'))

@app.route('/mobile-sao-do', methods=['GET'])
def mobile_sao_do():
    if session.get('role') != 'Sao đỏ': return redirect(url_for('login'))
    
    # [TÍNH NĂNG MỚI]: Bắt sóng Thông báo điều hành
    announcement_text = "🔔 Chào mừng các bạn Đội cờ đỏ! Chúc các bạn một tuần làm việc công tâm và trách nhiệm."
    import os, json
    config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config", "announcement.json")
    if os.path.exists(config_path):
        try:
            with open(config_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                announcement_text = data.get("text", announcement_text)
        except: pass
    try:
        with session_scope() as db_session:
            active_year = db_session.query(SchoolYear).filter_by(is_active=True).first()
            if not active_year: return "Hệ thống chưa mở năm học mới."

            # Dịch ngược từ Username (VD: SD15) ra ID của Sao đỏ (15)
            username = session.get('username', '')
            try: star_id = int(username.replace('SD', '').strip())
            except: return "Tài khoản không hợp lệ."

            assignment = db_session.query(Assignment).filter_by(red_star_id=star_id).order_by(Assignment.week_number.desc()).first()
            all_branches = db_session.query(Branch).filter_by(school_year_id=active_year.id).all()
            
            all_branches.sort(key=lambda b: [int(t) if t.isdigit() else t.lower() for t in re.split(r'(\d+)', str(b.name))])
            
            if not assignment:
                return render_template('sao_do_dashboard.html', assignment=None, all_branches=all_branches)
                
            current_week = f"Tuần {assignment.week_number}"
            
            # --- [VÁ LỖI AN TOÀN]: LẤY DANH SÁCH ĐỒNG ĐỘI CÙNG CA TRỰC ---
            teammates = []
            if assignment.duty_area:
                # Lấy tất cả người trực cùng ca, cùng tuần
                all_shift_assigns = db_session.query(Assignment).filter(
                    Assignment.week_number == assignment.week_number,
                    Assignment.shift == assignment.shift
                ).all()
                
                # Lọc ra những ai có trùng Object DutyArea và loại bản thân ra
                teammates = [a for a in all_shift_assigns if a.duty_area and a.duty_area.id == assignment.duty_area.id and a.red_star_id != star_id]
            # -----------------------------------------------------------
            
            # =========================================================
            # THUẬT TOÁN MỚI: TÍNH TOÁN NGÀY BẮT ĐẦU VÀ KẾT THÚC TUẦN
            # =========================================================
            start_date_str = ""
            end_date_str = ""
            
            existing_score = db_session.query(WeeklyScore).join(Branch).filter(
                WeeklyScore.week == current_week, 
                Branch.school_year_id == active_year.id
            ).first()
            
            if existing_score and existing_score.start_date: 
                start_date_str = existing_score.start_date
                end_date_str = existing_score.end_date or ""
            else:
                try:
                    import datetime as dt
                    if assignment and hasattr(assignment, 'date') and assignment.date:
                        py_date = assignment.date
                        start_date_str = py_date.strftime("%Y-%m-%d")
                        day_of_week = py_date.weekday()
                        if day_of_week <= 5: days_to_add = 5 - day_of_week
                        else: days_to_add = 6
                        end_date = py_date + dt.timedelta(days=days_to_add)
                        end_date_str = end_date.strftime("%Y-%m-%d")
                except Exception as e: pass

            if start_date_str and "-" in start_date_str:
                try: p = start_date_str.split('-'); start_date_str = f"{p[2]}/{p[1]}/{p[0]}"
                except: pass
            if end_date_str and "-" in end_date_str:
                try: p = end_date_str.split('-'); end_date_str = f"{p[2]}/{p[1]}/{p[0]}"
                except: pass
                
            if not start_date_str: start_date_str = "..."
            if not end_date_str: end_date_str = "..."
            # =========================================================

            target_classes = []
            if assignment.duty_area:
                import json
                import os
                config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config", "class_zones.json")
                if os.path.exists(config_path):
                    with open(config_path, "r", encoding="utf-8") as f:
                        zones_map = json.load(f)
                        raw_classes = zones_map.get(assignment.duty_area.name, [])
                        class_names = [str(c).strip().upper() for c in raw_classes]
                        
                        target_classes = db_session.query(Branch).filter(
                            Branch.name.in_(class_names), Branch.school_year_id == active_year.id
                        ).all()

            violations_bank = db_session.query(ViolationCategory).filter_by(school_year_id=active_year.id).all()
            
            existing_scores = {}
            scores_data_for_js = {} # [NÂNG CẤP]: Biến lưu trữ dữ liệu Ghi chú để đẩy ra JS
            
            for b in target_classes:
                sc = db_session.query(WeeklyScore).filter_by(branch_id=b.id, week=current_week).first()
                if sc: 
                    existing_scores[b.id] = sc
                    scores_data_for_js[b.id] = {
                        'note': sc.note or "",
                        'total_score': float(sc.total_score) if sc.total_score is not None else 100.0,
                        'score_tru': float(sc.score_tru) if sc.score_tru is not None else 0.0
                    }
                else:
                    scores_data_for_js[b.id] = {
                        'note': "",
                        'total_score': 100.0,
                        'score_tru': 0.0
                    }

            import json
            scores_data_json = json.dumps(scores_data_for_js)

            return render_template('sao_do_dashboard.html', 
                            assignment=assignment, 
                            target_classes=target_classes,
                            all_branches=all_branches,
                            violations_bank=violations_bank,
                            existing_scores=existing_scores,
                            scores_data_json=scores_data_json,
                            current_week=current_week,
                            teammates=teammates,
                            start_date=start_date_str,
                            end_date=end_date_str,
                            announcement_text=announcement_text
                        )
    except Exception as e:
        return f"Lỗi hệ thống Mobile: {e}"


# ==========================================
# MODULE: TRỢ LÝ AI PHÂN TÍCH VÀ VIẾT BÁO CÁO TUẦN (ĐÃ NÂNG CẤP CHUẨN QUẢN LÝ)
# ==========================================
@app.route('/api/ai_weekly_report/<week_name>')
def api_ai_weekly_report(week_name):
    try:
        with session_scope() as db_session:
            active_year = db_session.query(SchoolYear).filter_by(is_active=True).first()
            if not active_year: return {"error": "Chưa có năm học kích hoạt!"}
            
            # 1. Lấy toàn bộ điểm số của tuần
            scores = db_session.query(WeeklyScore).join(Branch).filter(
                WeeklyScore.week == week_name,
                Branch.school_year_id == active_year.id
            ).all()
            
            if not scores:
                return {"error": f"Chưa có dữ liệu điểm của {week_name} để phân tích!"}
                
            # 2. Tổng hợp dữ liệu thô để đưa cho AI
            total_classes = len(scores)
            avg_school_score = sum([float(sc.total_score or 0) for sc in scores]) / total_classes if total_classes > 0 else 0
            total_penalty_school = sum([float(sc.score_tru or 0) for sc in scores])
            
            top_classes = sorted(scores, key=lambda x: float(x.total_score or 0), reverse=True)[:3]
            bottom_classes = sorted(scores, key=lambda x: float(x.total_score or 0))[:3]
            
            # Tìm lớp bị trừ điểm nhiều nhất (Điểm nóng)
            worst_class_score = min(scores, key=lambda x: float(x.total_score or 0))
            
            summary_lines = []
            for sc in scores:
                b_name = sc.branch.name
                tot = sc.total_score
                rating = sc.week_rating
                note = sc.note or "Không có"
                summary_lines.append(f"- Lớp {b_name}: Tổng điểm {tot}, Xếp loại {rating}, Ghi chú: {note}")
                
            data_context = "\n".join(summary_lines)
            
            # 3. Đọc cấu hình Groq API Key từ file config
            api_key = ""
            config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config", "groq_settings.json")
            if os.path.exists(config_path):
                with open(config_path, "r", encoding="utf-8") as f:
                    cfg = json.load(f)
                    api_key = cfg.get("api_key", "")
                    
            if not api_key:
                return {"error": "Chưa cấu hình Groq API Key trong hệ thống!"}
                
            # 4. Gọi API Groq AI với Prompt chuẩn phong cách Nhà quản lý giáo dục
            import requests
            url = "https://api.groq.com/openai/v1/chat/completions"
            headers = {
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json"
            }
            
            prompt = f"""
            Bạn là một Chuyên gia Quản lý Giáo dục và Cố vấn Cấp cao cho Ban Giám hiệu trường THPT Thanh Hòa. Dựa trên dữ liệu tổng kết nề nếp của {week_name} dưới đây, hãy đưa ra bản phân tích mang tầm nhìn chiến lược, khách quan, sắc sảo và mang tính xây dựng cao:
            
            DỮ LIỆU THI ĐUA:
            - Tổng số lớp tham gia: {total_classes}
            - Điểm trung bình toàn trường: {avg_school_score:.1f} điểm
            - Tổng mức điểm trừ kỷ luật toàn trường: {total_penalty_school}đ
            - Lớp thấp điểm nhất (Điểm nóng): Lớp {worst_class_score.branch.name} (GVCN: {worst_class_score.branch.gvcn or 'Chưa cập nhật'}, Tổng điểm: {worst_class_score.total_score}, Lỗi: {worst_class_score.note})
            
            CHI TIẾT CÁC LỚP:
            {data_context}
            
            YÊU CẦU TRÌNH BÀY (BẮT BUỘC TRẢ VỀ ĐỊNH DẠNG HTML SẠCH SẼ):
            Hãy chia nội dung thành đúng 3 phần với các thẻ HTML sau (tuyệt đối không dùng dấu ** hay ký tự Markdown thô):
            
            1. Phần Bức tranh tổng quan (Tiêu đề dùng icon fa-chart-line): Nhận xét khái quát về biên độ điểm số, ý thức kỷ luật chung của học sinh toàn trường trong tuần.
            2. Phần Điểm nóng cần lưu ý (Tiêu đề dùng icon fa-triangle-exclamation): Chỉ ra tập thể lớp đang gặp vấn đề trầm trọng về nề nếp (ví dụ lớp {worst_class_score.branch.name}), phân tích nguyên nhân sơ bộ từ dữ liệu lỗi.
            3. Phần Đề xuất hướng giải quyết (Tiêu đề dùng icon fa-circle-check): Đưa ra các mốc giải pháp cụ thể dành cho Ban Giám hiệu, Đoàn trường phối hợp với Giáo viên chủ nhiệm để chấn chỉnh kỷ kỷ luật và duy trì phong trào.
            """
            
            payload = {
                "model": "openai/gpt-oss-120b",
                "messages": [{"role": "user", "content": prompt}],
                "temperature": 0.7
            }
            
            response = requests.post(url, json=payload, headers=headers, timeout=30)
            if response.status_code == 200:
                res_json = response.json()
                ai_text = res_json['choices'][0]['message']['content']
                
                # [THUẬT TOÁN LỌC VÀ XÓA DẤU SAO]: 
                
                # 1. Chuyển đổi định dạng **chữ đậm** thành thẻ <strong> của HTML
                ai_text = re.sub(r'\*\*(.*?)\*\*', r'<strong style="color: #0f172a;">\1</strong>', ai_text)
                # 2. Xóa bỏ hoàn toàn bất kỳ dấu sao đơn lẻ nào còn sót lại
                ai_text = ai_text.replace('*', '')
                
                # Xử lý ký tự xuống dòng an toàn tránh lỗi backslash
                replaced_text = ai_text.replace('\n', '<br>')
                
                if not ai_text.startswith("<div"):
                    formatted_html = f"""
                    <div style="font-family: 'Inter', sans-serif; color: #1e293b; line-height: 1.6;">
                        <div style="font-size: 14.5px; text-align: justify;">{replaced_text}</div>
                    </div>
                    """
                else:
                    formatted_html = ai_text
                
                return {"success": True, "report": formatted_html}
            else:
                return {"error": f"Lỗi kết nối AI API: {response.text}"}
                
    except Exception as e:
        return {"error": str(e)}
# ==========================================
# MODULE: TRA CỨU SỔ ĐEN TOÀN TRƯỜNG (ĐẦY ĐỦ MỌI BỘ LỌC)
# ==========================================
@app.route('/blacklist', methods=['GET'])
def blacklist():
    if session.get('role') not in ['Quản trị viên', 'Admin', 'Bí thư', 'Bí thư Đoàn trường']:
        flash("Bạn không có quyền truy cập Sổ đen!", "error")
        return redirect(url_for('dashboard'))
        
    try:
        with session_scope() as db_session:
            active_year = db_session.query(SchoolYear).filter_by(is_active=True).first()
            if not active_year:
                flash("Chưa có năm học kích hoạt!", "error")
                return redirect(url_for('dashboard'))
                
            branches = db_session.query(Branch).filter_by(school_year_id=active_year.id).all()
            categories = db_session.query(ViolationCategory).filter_by(school_year_id=active_year.id).all()
            
            # Lấy danh sách Tuần, Tháng, Học kỳ có dữ liệu trong CSDL để đổ vào bộ lọc thời gian
            weeks_db = db_session.query(WeeklyScore.week).join(Branch).filter(Branch.school_year_id == active_year.id).distinct().all()
            available_weeks = sorted([w[0] for w in weeks_db], key=lambda x: int(re.search(r'\d+', x).group()) if re.search(r'\d+', x) else 0)
            
            months_db = db_session.query(MonthlyRecord.month_name).filter(
                MonthlyRecord.school_year_id == active_year.id,
                MonthlyRecord.month_name.like('Tháng%')
            ).distinct().all()
            school_order = ["Tháng 9", "Tháng 10", "Tháng 11", "Tháng 12", "Tháng 1", "Tháng 2", "Tháng 3", "Tháng 4", "Tháng 5"]
            available_months = sorted([m[0] for m in months_db if m[0]], key=lambda x: school_order.index(x) if x in school_order else 99)
            
            semesters_db = db_session.query(MonthlyRecord.month_name).filter(
                MonthlyRecord.school_year_id == active_year.id,
                MonthlyRecord.month_name.like('Học kỳ%')
            ).distinct().all()
            available_semesters = [s[0] for s in semesters_db if s[0]]

            # Lấy toàn bộ tham số lọc từ giao diện người dùng gửi lên
            search_name = request.args.get('search_name', '').strip()
            search_branch = request.args.get('search_branch', '')
            search_violation = request.args.get('search_violation', '')
            time_mode = request.args.get('time_mode', 'all') # 'all', 'year', 'week', 'month', 'semester'
            time_value = request.args.get('time_value', '')

            query = db_session.query(
                WeeklyViolation, WeeklyScore, Branch, ViolationCategory
            ).join(WeeklyScore, WeeklyViolation.weekly_score_id == WeeklyScore.id)\
             .join(Branch, WeeklyScore.branch_id == Branch.id)\
             .join(ViolationCategory, WeeklyViolation.violation_id == ViolationCategory.id)\
             .filter(
                Branch.school_year_id == active_year.id,
                WeeklyViolation.student_name != None,
                WeeklyViolation.student_name != ''
             )
            
            # 1. Lọc theo Chi đoàn (Lớp)
            if search_branch and search_branch.isdigit():
                query = query.filter(Branch.id == int(search_branch))
                
            # 2. Lọc theo Lỗi vi phạm
            if search_violation and search_violation.isdigit():
                query = query.filter(ViolationCategory.id == int(search_violation))

            # 3. Lọc theo Thời gian (Tuần, Tháng, Học kỳ, Năm học)
            if time_mode == 'year':
                pass
            elif time_mode == 'week' and time_value:
                query = query.filter(WeeklyScore.week == time_value)
            elif time_mode in ['month', 'semester'] and time_value:
                m_rec = db_session.query(MonthlyRecord).filter_by(
                    school_year_id=active_year.id,
                    month_name=time_value
                ).first()
                if m_rec and m_rec.weeks_used:
                    valid_weeks = [w.strip() for w in m_rec.weeks_used.split(',') if w.strip()]
                    query = query.filter(WeeklyScore.week.in_(valid_weeks))
                else:
                    query = query.filter(WeeklyScore.week == 'NONE')

            results = query.order_by(WeeklyScore.id.desc(), Branch.name).all()
            
            violation_data = []
            for v, sc, b, c in results:
                # [THUẬT TOÁN ĐỒNG BỘ]: Bóc tách và chia đều lỗi học sinh dính chùm
                raw_names = str(v.student_name).replace(';', ',').split(',')
                valid_names = [n.strip().title() for n in raw_names if n.strip()]
                num_names = len(valid_names)
                qty_per_student = max(1, v.quantity // num_names) if num_names > 0 else v.quantity
                
                for n_clean in valid_names:
                    # 4. Lọc theo Tên học sinh (nếu có nhập)
                    if search_name and search_name.lower() not in n_clean.lower():
                        continue
                        
                    violation_data.append({
                        'week': sc.week,
                        'branch_name': b.name,
                        'student_name': n_clean,
                        'violation_name': c.name,
                        'quantity': qty_per_student, # <--- Đã sửa: Dùng số lượng đã chia đều
                        'penalty': float(c.penalty_points * qty_per_student) if getattr(c, 'point_type', 'Điểm trừ') != 'Điểm cộng' else 0
                    })
                
            return render_template('blacklist.html', 
                                   branches=branches, 
                                   categories=categories,
                                   available_weeks=available_weeks,
                                   available_months=available_months,
                                   available_semesters=available_semesters,
                                   violations=violation_data,
                                   search_name=search_name,
                                   search_branch=search_branch,
                                   search_violation=search_violation,
                                   time_mode=time_mode,
                                   time_value=time_value,
                                   active_year=active_year)
    except Exception as e:
        flash(f"Lỗi tải sổ đen: {e}", "error")
        return redirect(url_for('dashboard'))

# ==========================================
# API: XUẤT EXCEL SỔ ĐEN TOÀN TRƯỜNG (ÁP DỤNG ĐẦY ĐỦ BỘ LỌC)
# ==========================================
@app.route('/export_global_blacklist')
def export_global_blacklist():
    if session.get('role') not in ['Quản trị viên', 'Admin', 'Bí thư', 'Bí thư Đoàn trường']:
        flash("Bạn không có quyền xuất Sổ đen!", "error")
        return redirect(url_for('dashboard'))
        
    try:
        with session_scope() as db_session:
            active_year = db_session.query(SchoolYear).filter_by(is_active=True).first()
            if not active_year:
                return redirect(url_for('dashboard'))
                
            search_name = request.args.get('search_name', '').strip()
            search_branch = request.args.get('search_branch', '')
            search_violation = request.args.get('search_violation', '')
            time_mode = request.args.get('time_mode', 'all')
            time_value = request.args.get('time_value', '')
            
            query = db_session.query(
                WeeklyViolation, WeeklyScore, Branch, ViolationCategory
            ).join(WeeklyScore, WeeklyViolation.weekly_score_id == WeeklyScore.id)\
             .join(Branch, WeeklyScore.branch_id == Branch.id)\
             .join(ViolationCategory, WeeklyViolation.violation_id == ViolationCategory.id)\
             .filter(
                Branch.school_year_id == active_year.id,
                WeeklyViolation.student_name != None,
                WeeklyViolation.student_name != ''
             )
            
            if search_branch and search_branch.isdigit(): 
                query = query.filter(Branch.id == int(search_branch))
            if search_violation and search_violation.isdigit(): 
                query = query.filter(ViolationCategory.id == int(search_violation))

            if time_mode == 'week' and time_value:
                query = query.filter(WeeklyScore.week == time_value)
            elif time_mode in ['month', 'semester'] and time_value:
                m_rec = db_session.query(MonthlyRecord).filter_by(
                    school_year_id=active_year.id,
                    month_name=time_value
                ).first()
                if m_rec and m_rec.weeks_used:
                    valid_weeks = [w.strip() for w in m_rec.weeks_used.split(',') if w.strip()]
                    query = query.filter(WeeklyScore.week.in_(valid_weeks))
                else:
                    query = query.filter(WeeklyScore.week == 'NONE')
                
            results = query.order_by(Branch.name, WeeklyScore.id.desc()).all()
            
            violation_data = []
            filter_info = []
            
            for v, sc, b, c in results:
                # [THUẬT TOÁN ĐỒNG BỘ]: Bóc tách và chia đều lỗi học sinh dính chùm
                raw_names = str(v.student_name).replace(';', ',').split(',')
                valid_names = [n.strip().title() for n in raw_names if n.strip()]
                num_names = len(valid_names)
                qty_per_student = max(1, v.quantity // num_names) if num_names > 0 else v.quantity
                
                for n_clean in valid_names:
                    if search_name and search_name.lower() not in n_clean.lower(): continue
                    violation_data.append({
                        'week': sc.week,
                        'branch_name': b.name,
                        'student_name': n_clean,
                        'violation_name': c.name,
                        'quantity': qty_per_student # <--- Đã sửa: Dùng số lượng đã chia đều
                    })
                        
            if search_branch and search_branch.isdigit():
                b_obj = db_session.query(Branch).filter_by(id=int(search_branch)).first()
                if b_obj: filter_info.append(f"Lớp: {b_obj.name}")
            if search_violation and search_violation.isdigit():
                c_obj = db_session.query(ViolationCategory).filter_by(id=int(search_violation)).first()
                if c_obj: filter_info.append(f"Lỗi: {c_obj.name}")
            if search_name: filter_info.append(f"Tên HS: {search_name}")
            if time_mode == 'year': filter_info.append(f"Năm học: {active_year.name}")
            elif time_mode != 'all' and time_value: filter_info.append(f"Thời gian: {time_value}")

            import openpyxl
            from openpyxl.styles import Font, Alignment, Border, Side
            import io
            from flask import send_file
            
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "So_Den_Thong_Ke"
            
            ws.merge_cells('A1:F1')
            ws['A1'] = "ĐOÀN TRƯỜNG THPT THANH HÒA"
            ws['A1'].font = Font(name="Times New Roman", size=11, bold=True)
            ws.merge_cells('A3:F3')
            ws['A3'] = "THỐNG KÊ DANH SÁCH SỔ ĐEN KỶ LUẬT"
            ws['A3'].font = Font(name="Times New Roman", size=14, bold=True)
            ws['A3'].alignment = Alignment(horizontal="center")
            
            if filter_info:
                ws.merge_cells('A4:F4')
                ws['A4'] = f"Tiêu chí lọc: {', '.join(filter_info)}"
                ws['A4'].font = Font(name="Times New Roman", size=12, italic=True)
                ws['A4'].alignment = Alignment(horizontal="center")
            
            headers = ["STT", "Thời gian", "Chi đoàn", "Họ và Tên", "Lỗi Vi Phạm", "Số Lần"]
            thin = Side(border_style="thin", color="000000")
            border = Border(left=thin, right=thin, top=thin, bottom=thin)
            
            row_start = 6 if filter_info else 5
            for col, h in enumerate(headers, 1):
                c = ws.cell(row=row_start, column=col, value=h)
                c.font = Font(name="Times New Roman", size=11, bold=True)
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.border = border
                
            for idx, item in enumerate(violation_data, 1):
                row_idx = idx + row_start
                c1 = ws.cell(row=row_idx, column=1, value=idx)
                c2 = ws.cell(row=row_idx, column=2, value=item['week'])
                c3 = ws.cell(row=row_idx, column=3, value=item['branch_name'])
                c4 = ws.cell(row=row_idx, column=4, value=item['student_name'])
                c5 = ws.cell(row=row_idx, column=5, value=item['violation_name'])
                c6 = ws.cell(row=row_idx, column=6, value=item['quantity'])
                
                for cell in [c1, c2, c3, c4, c5, c6]:
                    cell.font = Font(name="Times New Roman", size=11)
                    cell.border = border
                c1.alignment = Alignment(horizontal="center")
                c2.alignment = Alignment(horizontal="center")
                c3.alignment = Alignment(horizontal="center")
                c6.alignment = Alignment(horizontal="center")
                
            ws.column_dimensions['A'].width = 6
            ws.column_dimensions['B'].width = 12
            ws.column_dimensions['C'].width = 12
            ws.column_dimensions['D'].width = 25
            ws.column_dimensions['E'].width = 35
            ws.column_dimensions['F'].width = 10
            
            log_system_action("XUẤT EXCEL", "Xuất Thống kê Sổ đen toàn trường")
            out = io.BytesIO()
            wb.save(out)
            out.seek(0)
            return send_file(out, download_name="Thong_Ke_So_Den.xlsx", as_attachment=True)
            
    except Exception as e:
        import traceback; traceback.print_exc()
        flash(f"Lỗi xuất Excel: {e}", "error")
        return redirect(url_for('blacklist'))
        
from flask import send_file

@app.route('/manifest.json')
def serve_manifest():
    return send_file('static/manifest.json', mimetype='application/manifest+json')

@app.route('/sw.js')
def serve_sw():
    return send_file('static/sw.js', mimetype='application/javascript')
from flask import send_from_directory

# ==========================================
# CẤU HÌNH APP HÓA (PWA) CHO ĐIỆN THOẠI
# ==========================================
@app.route('/sw.js')
def service_worker():
    # Phục vụ tệp sw.js từ thư mục static ra thẳng thư mục gốc mà không bị redirect
    return send_from_directory('static', 'sw.js', mimetype='application/javascript')

@app.route('/manifest.json')
def manifest():
    # Tương tự với file cấu hình PWA
    return send_from_directory('static', 'manifest.json', mimetype='application/json')

@app.route('/sao_do_quick_submit_form', methods=['POST'])
def sao_do_quick_submit_form():
    if session.get('role') != 'Sao đỏ': return redirect(url_for('login'))
    
    # ====================================================================
    # [KHIÊN BẢO VỆ]: CHỐNG NHÂN ĐÔI DỮ LIỆU DO TRÌNH DUYỆT TỰ ĐỘNG RETRY KHI RỚT MẠNG
    # ====================================================================
    import hashlib, time
    req_data = str(request.form.to_dict()) + str(request.form.get('evidence_base64', '')[:50])
    req_hash = hashlib.md5(req_data.encode('utf-8')).hexdigest()
    
    last_hash = session.get('last_quick_submit_hash')
    last_time = session.get('last_quick_submit_time', 0)
    current_time = time.time()
    
    if req_hash == last_hash and (current_time - last_time < 60):
        # Trả về thành công giả để trình duyệt ngừng gửi lại
        flash(f"⚡ Đã ghi nhận lỗi vào Sổ đen thành công!", "success")
        return redirect(url_for('mobile_sao_do'))
        
    session['last_quick_submit_hash'] = req_hash
    session['last_quick_submit_time'] = current_time
    # ====================================================================

    try:
        with session_scope() as db_session:
            active_year = db_session.query(SchoolYear).filter_by(is_active=True).first()
            week_name = request.form.get('week_name')
            raw_branch_id = request.form.get('branch_id')
            raw_viol_id = request.form.get('violation_id')
            student_name = request.form.get('student_name', '').strip()
            evidence_base64 = request.form.get('evidence_base64') # <--- Bổ sung dòng này            
            if not all([week_name, raw_branch_id, raw_viol_id]):
                flash("Lỗi: Vui lòng chọn đầy đủ Lớp và Lỗi vi phạm!", "error")
                return redirect(url_for('mobile_sao_do'))
                
            branch = db_session.query(Branch).filter_by(id=int(raw_branch_id)).first()
            violation = db_session.query(ViolationCategory).filter_by(id=int(raw_viol_id)).first()
            if not branch or not violation:
                return redirect(url_for('mobile_sao_do'))
                
            score = db_session.query(WeeklyScore).filter_by(branch_id=branch.id, week=week_name).first()
            # --- [BẢN VÁ LỖI 2]: KIỂM TRA KHÓA SỔ ---
            if score and getattr(score, 'is_locked', False):
                flash(f"⛔ Tuần {week_name} đã khóa sổ! Bạn không thể ghi nhận thêm lỗi.", "error")
                return redirect(url_for('mobile_sao_do'))
            if not score:
                score = WeeklyScore(branch_id=branch.id, week=week_name, week_rating='Bình thường', count_8=0, count_9=0, count_10=0, score_truc=100.0, score_cong=0.0, score_tru=0.0, note='', total_score=100.0)
                db_session.add(score)
                db_session.flush() 
            
            # --- [BỔ SUNG]: LƯU ẢNH MINH CHỨNG VÀ CHỐNG TRÙNG LẶP DO MẠNG YẾU ---
            # Lưu ý: Ở hàm submit_mobile_sao_do, thầy đổi biến week_name thành current_week cho khớp nhé
            saved_image_path = process_and_save_evidence(evidence_base64, branch.id, week_name) 
            if saved_image_path:
                current_images = getattr(score, 'evidence_image', '') or ''
                
                # Tách các link ảnh hiện có thành danh sách để rà soát
                existing_urls = [url.strip() for url in current_images.split('|') if url.strip()]
                new_urls = [url.strip() for url in saved_image_path.split('|') if url.strip()]
                
                # Chỉ ghép thêm đường link NẾU đường link đó chưa hề tồn tại trong CSDL
                for n_url in new_urls:
                    if n_url not in existing_urls:
                        existing_urls.append(n_url)
                        
                # Đóng gói lại thành chuỗi phân cách bằng dấu |
                score.evidence_image = "|".join(existing_urls)
                
            old_note = score.note if score and score.note else ""
            
            # --- [BẢN VÁ LỖI MÚI GIỜ]: TỰ ĐỘNG LẤY THỨ HIỆN TẠI (CHUẨN GIỜ VN) ---
            from datetime import datetime, timezone, timedelta
            vn_tz = timezone(timedelta(hours=7))
            days_vn = {0: '[T2]', 1: '[T3]', 2: '[T4]', 3: '[T5]', 4: '[T6]', 5: '[T7]', 6: '[CN]'}
            
            # Ép lấy giờ Việt Nam thay vì giờ máy chủ
            today_pfx = days_vn[datetime.now(vn_tz).weekday()]
            
            note_add = f"{today_pfx} {violation.name} x1 [{ ' '.join(student_name.split()).title() }]" if student_name else f"{today_pfx} {violation.name} x1"
            raw_combined_note = f"{old_note} ; {note_add}" if old_note else note_add
            
            all_categories = db_session.query(ViolationCategory).filter_by(school_year_id=active_year.id).all()
            sorted_cats = sorted(all_categories, key=lambda x: len(x.name), reverse=True)
            parsed_errors = {}
            
            for part in re.split(r'[,;+\n](?![^\[]*\])(?![^\(]*\))', raw_combined_note):
                part_clean = part.strip()
                if not part_clean: continue
                
                # --- [NÂNG CẤP]: BÓC TÁCH TAG NGÀY ĐỂ TRÁNH GỘP LỖI KHÁC NGÀY ---
                match_day = re.search(r'\[(T[2-7](?:\s*Chiều|\s*Chieu)?|CN)\]', part_clean, re.IGNORECASE)
                day_pfx = match_day.group(0) if match_day else ""
                text_to_parse = part_clean.replace(day_pfx, "").strip() if day_pfx else part_clean
                
                match_stu = re.search(r'\[(.*?)\]', text_to_parse)
                stu_name_raw = match_stu.group(1).strip() if match_stu else ""
                stu_name_normalized = " ".join(stu_name_raw.split()).title()
                
                matched = False
                for cat in sorted_cats:
                    if cat.name.lower() in text_to_parse.lower():
                        match_qty = re.search(r'(?:x|:|-)\s*(\d+)', text_to_parse.lower())
                        qty = int(match_qty.group(1)) if match_qty else 1
                        key = (cat.name, stu_name_normalized.lower(), stu_name_normalized, day_pfx) # Thêm day_pfx vào Key
                        parsed_errors[key] = parsed_errors.get(key, 0) + qty
                        matched = True
                        break
                if not matched:
                    parsed_errors[("MANUAL", text_to_parse.lower(), text_to_parse, day_pfx)] = 1
                    
            # --- XÉN TRẦN LỖI HỌC TẬP ---
            max_tot_bad = 14; max_mon_bad = 4
            settings = db_session.query(ScoreSettings).filter_by(school_year_id=active_year.id).first()
            if settings:
                max_tot_bad = int(getattr(settings, 'max_diem_tot', 14))
                max_mon_bad = int(getattr(settings, 'max_diem_mon', 4))
                
            bad_marks_expanded = []; other_errors = []
            for (cat_name, stu_key, stu_display, day_pfx), qty in parsed_errors.items():
                if cat_name == "MANUAL":
                    other_errors.append((cat_name, stu_display, qty, day_pfx))
                else:
                    is_bad_mark = "không học bài" in cat_name.lower() or "điểm kém" in cat_name.lower()
                    if is_bad_mark:
                        mon_match = re.search(r'\(Môn (.*?)\)', stu_display, re.IGNORECASE) if stu_display else None
                        mon = mon_match.group(1).strip() if mon_match else "Khác"
                        for _ in range(qty): bad_marks_expanded.append({'cat_name': cat_name, 'stu_display': stu_display, 'mon': mon, 'day_pfx': day_pfx})
                    else:
                        other_errors.append((cat_name, stu_display, qty, day_pfx))

            bad_by_subj = {}
            for bm in bad_marks_expanded:
                bad_by_subj.setdefault(bm['mon'], []).append(bm)
                
            surviving_bad_marks = []
            for marks in bad_by_subj.values(): surviving_bad_marks.extend(marks[:max_mon_bad])
            surviving_bad_marks = surviving_bad_marks[:max_tot_bad]
            
            capped_bad_counts = {}
            for bm in surviving_bad_marks:
                k = (bm['cat_name'], bm['stu_display'], bm['day_pfx'])
                capped_bad_counts[k] = capped_bad_counts.get(k, 0) + 1
                
            final_parts = []; diem_tru_auto = 0.0
            for (cat_name, stu_display, day_pfx), qty in capped_bad_counts.items():
                base_str = f"{cat_name} x{qty} [{stu_display}]" if stu_display else f"{cat_name} x{qty}"
                final_parts.append(f"{day_pfx} {base_str}".strip())
                for cat in sorted_cats:
                    if cat.name == cat_name and getattr(cat, 'point_type', 'Điểm trừ') != 'Điểm cộng':
                        diem_tru_auto += float(cat.penalty_points * qty); break
                        
            for cat_name, stu_display, qty, day_pfx in other_errors:
                if cat_name == "MANUAL": 
                    final_parts.append(f"{day_pfx} {stu_display}".strip() if day_pfx else stu_display)
                else:
                    base_str = f"{cat_name} x{qty} [{stu_display}]" if stu_display else f"{cat_name} x{qty}"
                    final_parts.append(f"{day_pfx} {base_str}".strip())
                    for cat in sorted_cats:
                        if cat.name == cat_name and getattr(cat, 'point_type', 'Điểm trừ') != 'Điểm cộng':
                            diem_tru_auto += float(cat.penalty_points * qty); break
                            
            score.note = " ; ".join(final_parts)
            score.score_tru = diem_tru_auto
            
            b_group = str(branch.group) if branch.group else "1"
            D8, D9, D10, TK, TT = 1.0, 3.0, 5.0, 20.0, 30.0
            if settings:
                D8, D9, D10 = float(settings.diem_8), float(settings.diem_9), float(settings.diem_10)
                TK, TT = float(settings.diem_tuan_kha), float(settings.diem_tuan_tot)
            
            diem_quy_uoc = (int(score.count_9 or 0)*D9) + (int(score.count_10 or 0)*D10)
            if "2" in b_group: diem_quy_uoc += (int(score.count_8 or 0)*D8)
            diem_xep_loai = TT if score.week_rating == "Tuần Tốt" else (TK if score.week_rating == "Tuần Khá" else 0.0)
            score.total_score = float(score.score_truc or 100.0) + diem_xep_loai + diem_quy_uoc + float(score.score_cong or 0.0) - diem_tru_auto

            db_session.query(WeeklyViolation).filter_by(weekly_score_id=score.id).delete()
            for (cat_name, stu_display, day_pfx), qty in capped_bad_counts.items():
                cat_id = next((c.id for c in sorted_cats if c.name == cat_name), None)
                if cat_id: db_session.add(WeeklyViolation(weekly_score_id=score.id, violation_id=cat_id, quantity=qty, student_name=stu_display if stu_display else None))
            
            for cat_name, stu_display, qty, day_pfx in other_errors:
                if cat_name != "MANUAL":
                    cat_id = next((c.id for c in sorted_cats if c.name == cat_name), None)
                    if cat_id: db_session.add(WeeklyViolation(weekly_score_id=score.id, violation_id=cat_id, quantity=qty, student_name=stu_display if stu_display else None))

            log_system_action("MOBILE TRỰC CỔNG", f"SD{session.get('username')} ghi nhận nhanh {branch.name}: {violation.name}")
            
            # =========================================================
            # [NÂNG CẤP]: BẮN THÔNG BÁO ĐẨY CHO GVCN NGAY KHI TRỰC CỔNG
            # =========================================================
            try:
                if note_add:
                    # Giả định tài khoản GVCN có định dạng "gvcn_10A1" 
                    # (Thầy/cô có thể sửa lại cho khớp với quy tắc đặt tên username GVCN của trường)
                    gvcn_username = branch.name.strip().upper() 
                    push_title = f"⚠️ Lớp {branch.name} vừa bị trừ điểm!"
                    push_body = f"Sao đỏ ghi nhận: {note_add[:50]}... \nBấm vào đây để xem chi tiết."
                    send_web_push(gvcn_username, push_title, push_body)
            except Exception as push_err:
                print(f"Lỗi gửi Push cho GVCN {branch.name}: {push_err}")
            # =========================================================

            flash(f"⚡ Đã ghi nhận lỗi của {branch.name} vào Sổ đen thành công!", "success")
            return redirect(url_for('mobile_sao_do'))
        
    except Exception as e:
        import traceback; traceback.print_exc()
        flash(f"Lỗi hệ thống: {e}", "error")
        return redirect(url_for('mobile_sao_do'))


@app.route('/submit_mobile_sao_do', methods=['POST'])
def submit_mobile_sao_do():
    if session.get('role') != 'Sao đỏ': return redirect(url_for('login'))
    
    # ====================================================================
    # [KHIÊN BẢO VỆ]: CHỐNG NHÂN ĐÔI DỮ LIỆU DO TRÌNH DUYỆT TỰ ĐỘNG RETRY KHI RỚT MẠNG
    # ====================================================================
    import hashlib, time
    req_data = str(request.form.to_dict()) + str(request.form.get('evidence_base64', '')[:50])
    req_hash = hashlib.md5(req_data.encode('utf-8')).hexdigest()
    
    last_hash = session.get('last_full_submit_hash')
    last_time = session.get('last_full_submit_time', 0)
    current_time = time.time()
    
    if req_hash == last_hash and (current_time - last_time < 60):
        flash(f"⚡ Hệ thống đã cập nhật điểm thành công!", "success")
        return redirect(url_for('mobile_sao_do'))
        
    session['last_full_submit_hash'] = req_hash
    session['last_full_submit_time'] = current_time
    # ====================================================================

    try:
        with session_scope() as db_session:
            active_year = db_session.query(SchoolYear).filter_by(is_active=True).first()
            current_week = request.form.get('week_name')
            branch_id = int(request.form.get('branch_id'))
            
            branch = db_session.query(Branch).filter_by(id=branch_id).first()
            score = db_session.query(WeeklyScore).filter_by(branch_id=branch.id, week=current_week).first()
            
            # --- KIỂM TRA KHÓA SỔ ---
            if score and getattr(score, 'is_locked', False):
                flash(f"⛔ Tuần {current_week} đã khóa sổ! Không thể sửa điểm của lớp {branch.name}.", "error")
                return redirect(url_for('mobile_sao_do'))
            # ----------------------------------------
            
            evidence_base64 = request.form.get('evidence_base64')
            
            settings = db_session.query(ScoreSettings).filter_by(school_year_id=active_year.id).first()
            max_mon = int(getattr(settings, 'max_diem_mon', 4)) if settings else 4
            max_tot = int(getattr(settings, 'max_diem_tot', 14)) if settings else 14
            
            rating = request.form.get('rating', score.week_rating if score else 'Bình thường')
            if rating == 'Bình thường' and score: rating = score.week_rating
            
            raw_scores_json = request.form.get('raw_scores_json', '').strip()
            f10, f9, f8 = 0, 0, 0
            b_group = str(branch.group) if branch.group else "1"
            has_new_raw_scores = False
            
            if raw_scores_json and raw_scores_json != '[]':
                try:
                    raw_list = json.loads(raw_scores_json)
                    if raw_list:
                        has_new_raw_scores = True
                        safe_week_key = f"{current_week}_Y{active_year.id}"
                        db_session.query(RawScore).filter_by(week=safe_week_key, branch_name=branch.name).delete()
                        remaining = max_tot
                        for item in raw_list:
                            subj = str(item.get("subj", "")).strip()
                            if not subj: continue
                            i10 = int(item.get("c10", 0) or 0); i9 = int(item.get("c9", 0) or 0); i8 = int(item.get("c8", 0) or 0)
                            if "1" in b_group: i8 = 0
                            db_session.add(RawScore(week=safe_week_key, branch_name=branch.name, subject=subj, c10=i10, c9=i9, c8=i8))
                            k10 = min(i10, max_mon); k9 = min(i9, max_mon - k10); k8 = min(i8, max_mon - k10 - k9)
                            a10 = min(k10, remaining); remaining -= a10
                            a9 = min(k9, remaining); remaining -= a9
                            a8 = min(k8, remaining); remaining -= a8
                            f10 += a10; f9 += a9; f8 += a8
                except Exception: pass
            
            if not has_new_raw_scores:
                f10 = score.count_10 if score else 0
                f9 = score.count_9 if score else 0
                f8 = score.count_8 if score else 0
                if "1" in b_group: f8 = 0

            # 2. HỢP NHẤT LỖI CŨ (TỪ DB) VÀ LỖI MỚI (TỪ APP)
            all_categories = db_session.query(ViolationCategory).filter_by(school_year_id=active_year.id).all()
            sorted_cats = sorted(all_categories, key=lambda x: len(x.name), reverse=True)
            
            # [BẢN VÁ LỖI MÚI GIỜ SAO ĐỎ]
            from datetime import datetime, timezone, timedelta
            vn_tz = timezone(timedelta(hours=7))
            days_vn = {0: '[T2]', 1: '[T3]', 2: '[T4]', 3: '[T5]', 4: '[T6]', 5: '[T7]', 6: '[CN]'}
            today_pfx = days_vn[datetime.now(vn_tz).weekday()]
            
            old_note = score.note if score and score.note else ""
            
            new_checkbox_notes = []
            for cat in all_categories:
                qty_new = int(request.form.get(f'viol_{cat.id}', '0'))
                if qty_new > 0:
                    stu_new = request.form.get(f'student_{cat.id}', '').strip()
                    stu_norm = " ".join(stu_new.split()).title() if stu_new else ""
                    if stu_norm:
                        new_checkbox_notes.append(f"{today_pfx} {cat.name} x{qty_new} [{stu_norm}]")
                    else:
                        new_checkbox_notes.append(f"{today_pfx} {cat.name} x{qty_new}")
                        
            raw_app_note = request.form.get('manual_note', '').strip()
            manual_parts = []
            if raw_app_note:
                diff_note = raw_app_note
                if old_note and raw_app_note.startswith(old_note):
                    diff_note = raw_app_note.replace(old_note, "", 1).strip().lstrip(" ;,")
                if diff_note:
                    for part in smart_split_note(diff_note):
                        p_clean = part.strip()
                        if p_clean:
                            if not re.search(r'\[(T[2-7]|CN)\]', p_clean):
                                manual_parts.append(f"{today_pfx} {p_clean}")
                            else:
                                manual_parts.append(p_clean)

            combined_parts = []
            if old_note: combined_parts.append(old_note)
            if new_checkbox_notes: combined_parts.extend(new_checkbox_notes)
            if manual_parts: combined_parts.extend(manual_parts)
            
            raw_combined_note = " ; ".join(combined_parts)

            parsed_errors = {}
            for part in smart_split_note(raw_combined_note):
                part_clean = part.strip()
                if not part_clean: continue
                
                match_day = re.search(r'\[(T[2-7](?:\s*Chiều|\s*Chieu)?|CN)\]', part_clean, re.IGNORECASE)
                day_pfx = match_day.group(0) if match_day else ""
                text_to_parse = part_clean.replace(day_pfx, "").strip() if day_pfx else part_clean
                
                match_stu = re.search(r'\[(.*?)\]', text_to_parse)
                stu_name_raw = match_stu.group(1).strip() if match_stu else ""
                stu_name_normalized = " ".join(stu_name_raw.split()).title()
                
                matched = False
                for cat in sorted_cats:
                    if cat.name.lower() in text_to_parse.lower():
                        match_qty = re.search(r'(?:x|:|-)\s*(\d+)', text_to_parse.lower())
                        qty = int(match_qty.group(1)) if match_qty else 1
                        key = (cat.name, stu_name_normalized.lower(), stu_name_normalized, day_pfx)
                        parsed_errors[key] = parsed_errors.get(key, 0) + qty
                        matched = True
                        break
                if not matched:
                    parsed_errors[("MANUAL", text_to_parse.lower(), text_to_parse, day_pfx)] = 1

            # 3. CHẠY THUẬT TOÁN XÉN TRẦN BAREM TRÊN TỔNG DỮ LIỆU ĐÃ HỢP NHẤT
            bad_marks_expanded = []
            other_errors = []

            for (cat_name, stu_key, stu_display, day_pfx), qty in parsed_errors.items():
                if cat_name == "MANUAL":
                    other_errors.append((cat_name, stu_display, qty, day_pfx))
                else:
                    is_bad_mark = "không học bài" in cat_name.lower() or "điểm kém" in cat_name.lower()
                    if is_bad_mark:
                        mon_match = re.search(r'\(Môn (.*?)\)', stu_display, re.IGNORECASE) if stu_display else None
                        mon = mon_match.group(1).strip() if mon_match else "Khác"
                        for _ in range(qty):
                            bad_marks_expanded.append({'cat_name': cat_name, 'stu_display': stu_display, 'mon': mon, 'day_pfx': day_pfx})
                    else:
                        other_errors.append((cat_name, stu_display, qty, day_pfx))

            bad_by_subj = {}
            for bm in bad_marks_expanded:
                bad_by_subj.setdefault(bm['mon'], []).append(bm)
                
            surviving_bad_marks = []
            for marks in bad_by_subj.values(): surviving_bad_marks.extend(marks[:max_mon])
            surviving_bad_marks = surviving_bad_marks[:max_tot]

            capped_bad_counts = {}
            for bm in surviving_bad_marks:
                k = (bm['cat_name'], bm['stu_display'], bm['day_pfx'])
                capped_bad_counts[k] = capped_bad_counts.get(k, 0) + 1

            # 4. TÍNH TỔNG ĐIỂM TRỪ VÀ TẠO CHUỖI GHI CHÚ MỚI
            diem_tru_final = 0.0
            final_note_parts = []

            for (cat_name, stu_display, day_pfx), qty in capped_bad_counts.items():
                base_str = f"{cat_name} x{qty} [{stu_display}]" if stu_display else f"{cat_name} x{qty}"
                final_note_parts.append(f"{day_pfx} {base_str}".strip())
                for cat in sorted_cats:
                    if cat.name == cat_name and getattr(cat, 'point_type', 'Điểm trừ') != 'Điểm cộng':
                        diem_tru_final += float(cat.penalty_points * qty); break
                        
            for cat_name, stu_display, qty, day_pfx in other_errors:
                if cat_name == "MANUAL":
                    final_note_parts.append(f"{day_pfx} {stu_display}".strip() if day_pfx else stu_display)
                else:
                    base_str = f"{cat_name} x{qty} [{stu_display}]" if stu_display else f"{cat_name} x{qty}"
                    final_note_parts.append(f"{day_pfx} {base_str}".strip())
                    for cat in sorted_cats:
                        if cat.name == cat_name and getattr(cat, 'point_type', 'Điểm trừ') != 'Điểm cộng':
                            diem_tru_final += float(cat.penalty_points * qty); break

            final_note = " ; ".join(final_note_parts)

            # 5. TÍNH TỔNG ĐIỂM DỰ KIẾN
            truc = float(score.score_truc) if score and score.score_truc is not None else 100.0
            cong = float(score.score_cong) if score and score.score_cong is not None else 0.0
            
            D8, D9, D10, TK, TT = 1.0, 3.0, 5.0, 20.0, 30.0
            if settings:
                D8, D9, D10 = float(settings.diem_8), float(settings.diem_9), float(settings.diem_10)
                TK, TT = float(settings.diem_tuan_kha), float(settings.diem_tuan_tot)
            
            diem_quy_uoc = (f9 * D9) + (f10 * D10)
            if "2" in b_group: diem_quy_uoc += (f8 * D8)

            diem_xep_loai = 0.0
            if rating == "Tuần Tốt": diem_xep_loai = TT
            elif rating == "Tuần Khá": diem_xep_loai = TK
            
            total_val = truc + diem_xep_loai + diem_quy_uoc + cong - diem_tru_final

            # 6. GHI VÀO DATABASE
            saved_image_path = process_and_save_evidence(evidence_base64, branch.id, current_week)
            
            if score:
                score.week_rating = rating
                score.count_8 = f8; score.count_9 = f9; score.count_10 = f10
                score.note = final_note
                score.score_tru = diem_tru_final
                score.total_score = total_val
                
                # --- [BẢN VÁ LỖI TỐI THƯỢNG]: LỌC ẢNH TRÙNG LẶP TRONG CSDL ---
                if saved_image_path:
                    current_images = getattr(score, 'evidence_image', '') or ''
                    existing_urls = [url.strip() for url in current_images.split('|') if url.strip()]
                    new_urls = [url.strip() for url in saved_image_path.split('|') if url.strip()]
                    
                    for n_url in new_urls:
                        if n_url not in existing_urls:
                            existing_urls.append(n_url)
                            
                    score.evidence_image = "|".join(existing_urls)
                # -------------------------------------------------------------
            else:
                score = WeeklyScore(
                    branch_id=branch.id, week=current_week, week_rating=rating,
                    count_8=f8, count_9=f9, count_10=f10,
                    score_truc=truc, score_cong=cong, score_tru=diem_tru_final,
                    note=final_note, total_score=total_val,
                    evidence_image=saved_image_path
                )
                db_session.add(score)
                db_session.flush()

            db_session.query(WeeklyViolation).filter_by(weekly_score_id=score.id).delete()
            
            for (cat_name, stu_display, day_pfx), qty in capped_bad_counts.items():
                cat_id = next((c.id for c in sorted_cats if c.name == cat_name), None)
                if cat_id: db_session.add(WeeklyViolation(weekly_score_id=score.id, violation_id=cat_id, quantity=qty, student_name=stu_display if stu_display else None))
            for cat_name, stu_display, qty, day_pfx in other_errors:
                if cat_name != "MANUAL":
                    cat_id = next((c.id for c in sorted_cats if c.name == cat_name), None)
                    if cat_id: db_session.add(WeeklyViolation(weekly_score_id=score.id, violation_id=cat_id, quantity=qty, student_name=stu_display if stu_display else None))

            log_system_action("MOBILE SAO ĐỎ", f"SD{session.get('username')} đã chấm điểm và cập nhật lớp {branch.name}.")
            
            # =========================================================
            # [NÂNG CẤP]: BẮN THÔNG BÁO ĐẨY CHO GVCN KHI CHỐT SỔ LỚP
            # =========================================================
            try:
                new_errors = []
                if new_checkbox_notes: new_errors.extend(new_checkbox_notes)
                if manual_parts: new_errors.extend(manual_parts)
                
                if new_errors:
                    error_summary = " ; ".join(new_errors)
                    gvcn_username = branch.name.strip().upper()
                    push_title = f"⚠️ Lớp {branch.name} vừa bị trừ điểm!"
                    push_body = f"Sao đỏ ghi nhận: {error_summary[:50]}... \nBấm vào đây để xem chi tiết."
                    send_web_push(gvcn_username, push_title, push_body)
            except Exception as push_err:
                print(f"Lỗi gửi Push cho GVCN {branch.name}: {push_err}")
            # =========================================================

            flash(f"Đã nộp điểm và đồng bộ vào hệ thống cho lớp {branch.name} thành công!", "success")
            return redirect(url_for('mobile_sao_do'))
            
    except Exception as e:
        import traceback; traceback.print_exc()
        flash(f"Lỗi: {e}", "error")
        return redirect(url_for('mobile_sao_do'))
    
# =====================================================================
# MODULE: PHIẾU PHÂN TÍCH CHUYÊN SÂU LỚP HỌC (DÀNH CHO HỌP GVCN)
# =====================================================================
@app.route('/class_monthly_analysis', methods=['GET', 'POST'])
def class_monthly_analysis():
    try:
        with session_scope() as db_session:
            active_year = db_session.query(SchoolYear).filter_by(is_active=True).first()
            if not active_year:
                flash("Chưa có năm học kích hoạt!", "error")
                return redirect(url_for('dashboard'))

            # Lấy danh sách tháng đã chốt, sắp xếp theo thứ tự năm học
            months_db = db_session.query(MonthlyRecord.month_name).filter(
                MonthlyRecord.school_year_id == active_year.id,
                MonthlyRecord.month_name.like('Tháng%')
            ).distinct().all()
            
            school_order = ["Tháng 9", "Tháng 10", "Tháng 11", "Tháng 12", "Tháng 1", "Tháng 2", "Tháng 3", "Tháng 4", "Tháng 5"]
            raw_months = [m[0] for m in months_db if m[0]]
            available_months = sorted(raw_months, key=lambda x: school_order.index(x) if x in school_order else 99)
            
            branches = db_session.query(Branch).filter_by(school_year_id=active_year.id).all()

            selected_month = request.args.get('month', available_months[-1] if available_months else "")
            selected_branch_id = request.args.get('branch_id', type=int)

            analysis_data = None

            if selected_month and selected_branch_id:
                branch = db_session.query(Branch).filter_by(id=selected_branch_id).first()
                month_rec = db_session.query(MonthlyRecord).filter_by(
                    school_year_id=active_year.id, 
                    month_name=selected_month, 
                    branch_id=selected_branch_id
                ).first()

                if month_rec and month_rec.weeks_used:
                    valid_weeks = [w.strip() for w in month_rec.weeks_used.split(',') if w.strip()]
                    
                    # 1. TRÍCH XUẤT DỮ LIỆU NỀ NẾP (KỶ LUẬT)
                    violations = db_session.query(WeeklyViolation, ViolationCategory).join(
                        ViolationCategory, WeeklyViolation.violation_id == ViolationCategory.id
                    ).join(
                        WeeklyScore, WeeklyViolation.weekly_score_id == WeeklyScore.id
                    ).filter(
                        WeeklyScore.week.in_(valid_weeks),
                        WeeklyScore.branch_id == selected_branch_id
                    ).all()

                    viol_summary = {}
                    student_issues = {}
                    student_bonus = {} 

                    for v, cat in violations:
                        qty = v.quantity or 1
                        # Gom nhóm lỗi
                        if getattr(cat, 'point_type', 'Điểm trừ') == 'Điểm trừ':
                            viol_summary[cat.name] = viol_summary.get(cat.name, 0) + qty
                            
                            # Gom nhóm cá nhân vi phạm
                            if v.student_name:
                                names = [n.strip().title() for n in str(v.student_name).replace(';', ',').split(',') if n.strip()]
                                for name in names:
                                    if name not in student_issues: student_issues[name] = []
                                    student_issues[name].append(f"{cat.name} (x{qty})")
                        else:
                            # Điểm cộng (Học sinh xuất sắc)
                            if v.student_name:
                                names = [n.strip().title() for n in str(v.student_name).replace(';', ',').split(',') if n.strip()]
                                for name in names:
                                    student_bonus[name] = student_bonus.get(name, 0) + qty

                    # [ĐÃ THÁO KHÓA]: Hiển thị TOÀN BỘ lỗi, sắp xếp từ nhiều đến ít (Không giới hạn Top 3 nữa)
                    all_violations_sorted = sorted(viol_summary.items(), key=lambda x: x[1], reverse=True)
                    
                    # [ĐÃ THÁO KHÓA]: Hiển thị TOÀN BỘ học sinh có tên trong sổ đen (Kể cả 1 lần)
                    all_violators_sorted = {k: v for k, v in sorted(student_issues.items(), key=lambda item: len(item[1]), reverse=True)}

                    # 2. TRÍCH XUẤT DỮ LIỆU HỌC TẬP (RAW SCORES)
                    raw_scores = db_session.query(RawScore).filter(
                        RawScore.week.in_(valid_weeks),
                        RawScore.branch_name == branch.name
                    ).all()

                    subject_scores = {}
                    total_good_points = 0
                    for rs in raw_scores:
                        subj = rs.subject.strip()
                        c10, c9, c8 = int(rs.c10 or 0), int(rs.c9 or 0), int(rs.c8 or 0)
                        if "1" in str(branch.group): c8 = 0
                        
                        points = c10 + c9 + c8
                        total_good_points += points
                        subject_scores[subj] = subject_scores.get(subj, 0) + points

                    # Chỉ lấy Top 3 môn tốt nhất, bỏ phần quét môn yếu kém
                    top_subjects = sorted(subject_scores.items(), key=lambda x: x[1], reverse=True)[:3]

                    # Thống kê Tuần Tốt / Khá
                    scores = db_session.query(WeeklyScore).filter(
                        WeeklyScore.week.in_(valid_weeks),
                        WeeklyScore.branch_id == selected_branch_id
                    ).all()
                    
                    tuan_tot = sum(1 for s in scores if s.week_rating == 'Tuần Tốt')
                    tuan_kha = sum(1 for s in scores if s.week_rating == 'Tuần Khá')

                    analysis_data = {
                        "branch": branch,
                        "month_name": selected_month,
                        "weeks_used": month_rec.weeks_used,
                        "total_score": month_rec.total_score,
                        "rank": month_rec.rank,
                        "top_violations": all_violations_sorted, 
                        "frequent_violators": all_violators_sorted, 
                        "student_bonus": sorted(student_bonus.items(), key=lambda x: x[1], reverse=True),
                        "total_good_points": total_good_points,
                        "top_subjects": top_subjects,
                        "tuan_tot": tuan_tot,
                        "tuan_kha": tuan_kha
                    }
            return render_template('class_monthly_analysis.html', 
                                   available_months=available_months, 
                                   branches=branches, 
                                   selected_month=selected_month, 
                                   selected_branch_id=selected_branch_id,
                                   data=analysis_data)
    except Exception as e:
        import traceback; traceback.print_exc()
        flash(f"Lỗi tải trang phân tích: {e}", "error")
        return redirect(url_for('dashboard'))
    
# =====================================================================
# MODULE: XUẤT PDF PHIẾU PHÂN TÍCH TOÀN BỘ CÁC LỚP
# =====================================================================
@app.route('/export_all_monthly_analysis', methods=['GET'])
def export_all_monthly_analysis():
    try:
        with session_scope() as db_session:
            active_year = db_session.query(SchoolYear).filter_by(is_active=True).first()
            if not active_year:
                return "Chưa có năm học kích hoạt!"

            selected_month = request.args.get('month')
            if not selected_month:
                return "Lỗi: Vui lòng chọn tháng trước khi xuất!"

            branches = db_session.query(Branch).filter_by(school_year_id=active_year.id).order_by(Branch.name).all()
            all_data = []

            for branch in branches:
                month_rec = db_session.query(MonthlyRecord).filter_by(
                    school_year_id=active_year.id, 
                    month_name=selected_month, 
                    branch_id=branch.id
                ).first()

                if month_rec and month_rec.weeks_used:
                    valid_weeks = [w.strip() for w in month_rec.weeks_used.split(',') if w.strip()]
                    
                    # NỀ NẾP
                    violations = db_session.query(WeeklyViolation, ViolationCategory).join(
                        ViolationCategory, WeeklyViolation.violation_id == ViolationCategory.id
                    ).join(
                        WeeklyScore, WeeklyViolation.weekly_score_id == WeeklyScore.id
                    ).filter(WeeklyScore.week.in_(valid_weeks), WeeklyScore.branch_id == branch.id).all()

                    viol_summary = {}; student_issues = {}; student_bonus = {} 
                    for v, cat in violations:
                        qty = v.quantity or 1
                        if getattr(cat, 'point_type', 'Điểm trừ') == 'Điểm trừ':
                            viol_summary[cat.name] = viol_summary.get(cat.name, 0) + qty
                            if v.student_name:
                                names = [n.strip().title() for n in str(v.student_name).replace(';', ',').split(',') if n.strip()]
                                for name in names:
                                    if name not in student_issues: student_issues[name] = []
                                    student_issues[name].append(f"{cat.name} (x{qty})")
                        else:
                            if v.student_name:
                                names = [n.strip().title() for n in str(v.student_name).replace(';', ',').split(',') if n.strip()]
                                for name in names: student_bonus[name] = student_bonus.get(name, 0) + qty

                    all_violations_sorted = sorted(viol_summary.items(), key=lambda x: x[1], reverse=True)
                    all_violators_sorted = {k: v for k, v in sorted(student_issues.items(), key=lambda item: len(item[1]), reverse=True)}

                    # HỌC TẬP
                    raw_scores = db_session.query(RawScore).filter(RawScore.week.in_(valid_weeks), RawScore.branch_name == branch.name).all()
                    subject_scores = {}; total_good_points = 0
                    for rs in raw_scores:
                        subj = rs.subject.strip()
                        c10, c9, c8 = int(rs.c10 or 0), int(rs.c9 or 0), int(rs.c8 or 0)
                        if "1" in str(branch.group): c8 = 0
                        points = c10 + c9 + c8
                        total_good_points += points
                        subject_scores[subj] = subject_scores.get(subj, 0) + points

                    top_subjects = sorted(subject_scores.items(), key=lambda x: x[1], reverse=True)[:3]
                    scores = db_session.query(WeeklyScore).filter(WeeklyScore.week.in_(valid_weeks), WeeklyScore.branch_id == branch.id).all()
                    tuan_tot = sum(1 for s in scores if s.week_rating == 'Tuần Tốt')
                    tuan_kha = sum(1 for s in scores if s.week_rating == 'Tuần Khá')

                    all_data.append({
                        "branch": branch, "month_name": selected_month, "weeks_used": month_rec.weeks_used,
                        "total_score": month_rec.total_score, "rank": month_rec.rank,
                        "top_violations": all_violations_sorted, "frequent_violators": all_violators_sorted, 
                        "student_bonus": sorted(student_bonus.items(), key=lambda x: x[1], reverse=True),
                        "total_good_points": total_good_points, "top_subjects": top_subjects,
                        "tuan_tot": tuan_tot, "tuan_kha": tuan_kha
                    })

            return render_template('class_monthly_analysis_all.html', all_data=all_data, selected_month=selected_month)
    except Exception as e:
        import traceback; traceback.print_exc()
        return f"<h1>Lỗi xuất PDF toàn trường: {e}</h1>"
# =====================================================================
# MODULE: PHIẾU PHÂN TÍCH ĐÁNH GIÁ CUỐI HỌC KỲ (HỌP GVCN)
# =====================================================================
@app.route('/class_semester_analysis', methods=['GET', 'POST'])
def class_semester_analysis():
    try:
        with session_scope() as db_session:
            active_year = db_session.query(SchoolYear).filter_by(is_active=True).first()
            if not active_year:
                flash("Chưa có năm học kích hoạt!", "error")
                return redirect(url_for('dashboard'))

            # [ĐÃ SỬA]: Quét danh sách Học kỳ từ bảng MonthlyRecord giống hệt logic của đồng chí
            semesters_db = db_session.query(MonthlyRecord.month_name).filter(
                MonthlyRecord.school_year_id == active_year.id,
                MonthlyRecord.month_name.like('Học kỳ%')
            ).distinct().all()
            
            available_semesters = [s[0] for s in semesters_db if s[0]]
            if not available_semesters: available_semesters = ["Học kỳ 1", "Học kỳ 2"]

            branches = db_session.query(Branch).filter_by(school_year_id=active_year.id).all()

            selected_semester = request.args.get('semester', available_semesters[0])
            selected_branch_id = request.args.get('branch_id', type=int)

            analysis_data = None

            if selected_semester and selected_branch_id:
                branch = db_session.query(Branch).filter_by(id=selected_branch_id).first()
                
                # [ĐÃ SỬA]: Truy vấn dữ liệu Học kỳ từ bảng MonthlyRecord
                semester_rec = db_session.query(MonthlyRecord).filter_by(
                    school_year_id=active_year.id, 
                    month_name=selected_semester, 
                    branch_id=selected_branch_id
                ).first()

                if semester_rec and semester_rec.weeks_used:
                    # Trong bảng này, weeks_used đang chứa chuỗi các tháng (VD: "Tháng 9, Tháng 10")
                    valid_months = [m.strip() for m in semester_rec.weeks_used.split(',') if m.strip()]
                    
                    # --- TIÊU CHÍ 1: KẾT QUẢ THI ĐUA CHUNG ---
                    monthly_records = db_session.query(MonthlyRecord).filter(
                        MonthlyRecord.school_year_id == active_year.id,
                        MonthlyRecord.branch_id == selected_branch_id,
                        MonthlyRecord.month_name.in_(valid_months)
                    ).all()
                    
                    school_order = ["Tháng 9", "Tháng 10", "Tháng 11", "Tháng 12", "Tháng 1", "Tháng 2", "Tháng 3", "Tháng 4", "Tháng 5"]
                    monthly_records = sorted(monthly_records, key=lambda x: school_order.index(x.month_name) if x.month_name in school_order else 99)
                    
                    monthly_stats = [{"month": m.month_name, "score": float(m.total_score), "rank": m.rank} for m in monthly_records]
                    avg_score = round(sum(m['score'] for m in monthly_stats) / len(monthly_stats), 2) if monthly_stats else 0
                    
                    progress_note = "Chưa đủ dữ liệu để so sánh"
                    if len(monthly_stats) >= 2:
                        first_m = monthly_stats[0]
                        last_m = monthly_stats[-1]
                        if last_m['rank'] < first_m['rank']:
                            progress_note = f"Tiến bộ rõ rệt (Từ Hạng {first_m['rank']} vươn lên Hạng {last_m['rank']})"
                        elif last_m['rank'] > first_m['rank']:
                            progress_note = f"Có dấu hiệu giảm sút (Từ Hạng {first_m['rank']} rớt xuống Hạng {last_m['rank']})"
                        else:
                            progress_note = f"Duy trì kết quả ổn định ở Hạng {last_m['rank']}"

                    valid_weeks = []
                    for m_rec in monthly_records:
                        if m_rec.weeks_used: valid_weeks.extend([w.strip() for w in m_rec.weeks_used.split(',') if w.strip()])
                    valid_weeks = list(set(valid_weeks))

                    # --- TIÊU CHÍ 2: NỀN NẾP & KỶ LUẬT ---
                    violations = db_session.query(WeeklyViolation, ViolationCategory).join(
                        ViolationCategory, WeeklyViolation.violation_id == ViolationCategory.id
                    ).join(
                        WeeklyScore, WeeklyViolation.weekly_score_id == WeeklyScore.id
                    ).filter(WeeklyScore.week.in_(valid_weeks), WeeklyScore.branch_id == selected_branch_id).all()

                    viol_summary = {}
                    student_issues = {}
                    total_violations = 0
                    
                    for v, cat in violations:
                        qty = v.quantity or 1
                        if getattr(cat, 'point_type', 'Điểm trừ') == 'Điểm trừ':
                            total_violations += qty
                            viol_summary[cat.name] = viol_summary.get(cat.name, 0) + qty
                            
                            if v.student_name:
                                names = [n.strip().title() for n in str(v.student_name).replace(';', ',').split(',') if n.strip()]
                                for name in names:
                                    if name not in student_issues: student_issues[name] = []
                                    student_issues[name].append(f"{cat.name} (x{qty})")

                    top_violations = sorted(viol_summary.items(), key=lambda x: x[1], reverse=True)
                    severe_violators = {k: v for k, v in sorted(student_issues.items(), key=lambda item: len(item[1]), reverse=True) if len(v) >= 3} 

                    analysis_data = {
                        "branch": branch,
                        "semester_name": selected_semester,
                        "months_used": semester_rec.weeks_used, 
                        "total_score": semester_rec.total_score,
                        "rank": semester_rec.rank,
                        "monthly_stats": monthly_stats,
                        "avg_score": avg_score,
                        "progress_note": progress_note,
                        "total_violations": total_violations,
                        "top_violations": top_violations,
                        "severe_violators": severe_violators
                    }

            return render_template('class_semester_analysis.html', 
                                   available_semesters=available_semesters, 
                                   branches=branches, 
                                   selected_semester=selected_semester, 
                                   selected_branch_id=selected_branch_id,
                                   data=analysis_data)
    except Exception as e:
        import traceback; traceback.print_exc()
        flash(f"Lỗi tải trang phân tích học kỳ: {e}", "error")
        return redirect(url_for('dashboard'))
    
# =====================================================================
# MODULE: BÁO CÁO TỔNG QUAN TOÀN TRƯỜNG (DÀNH CHO BÍ THƯ / HIỆU TRƯỞNG)
# =====================================================================
@app.route('/school_monthly_analysis', methods=['GET'])
def school_monthly_analysis():
    try:
        with session_scope() as db_session:
            active_year = db_session.query(SchoolYear).filter_by(is_active=True).first()
            if not active_year:
                flash("Chưa có năm học kích hoạt!", "error")
                return redirect(url_for('dashboard'))

            months_db = db_session.query(MonthlyRecord.month_name).filter(
                MonthlyRecord.school_year_id == active_year.id,
                MonthlyRecord.month_name.like('Tháng%')
            ).distinct().all()
            
            school_order = ["Tháng 9", "Tháng 10", "Tháng 11", "Tháng 12", "Tháng 1", "Tháng 2", "Tháng 3", "Tháng 4", "Tháng 5"]
            raw_months = [m[0] for m in months_db if m[0]]
            available_months = sorted(raw_months, key=lambda x: school_order.index(x) if x in school_order else 99)

            selected_month = request.args.get('month', available_months[-1] if available_months else "")
            analysis_data = None

            if selected_month:
                # 1. Lấy toàn bộ danh sách Chi đoàn của năm học hiện tại làm bản đồ tra cứu chuẩn tuyệt đối
                all_branches = db_session.query(Branch).filter_by(school_year_id=active_year.id).all()
                branch_map = {b.id: b for b in all_branches}
                branch_group_map = {b.name: b.group for b in all_branches}

                # 2. Lấy tất cả các bản ghi điểm tháng đã chốt
                records = db_session.query(MonthlyRecord).filter(
                    MonthlyRecord.school_year_id == active_year.id, 
                    MonthlyRecord.month_name == selected_month
                ).order_by(MonthlyRecord.rank).all()
                
                processed_records = []
                valid_weeks = set()
                
                for m_rec in records:
                    # Tra cứu thông tin lớp an toàn tuyệt đối qua branch_id
                    b_rec = branch_map.get(m_rec.branch_id)
                    
                    b_name = str(b_rec.name).strip() if b_rec and b_rec.name else f"Chi đoàn {m_rec.branch_id}"
                    b_gvcn = str(b_rec.gvcn).strip() if b_rec and b_rec.gvcn else "Chưa cập nhật GVCN"

                    processed_records.append({
                        "rank": m_rec.rank,
                        "total_score": round(float(m_rec.total_score or 0), 1),
                        "branch_name": b_name,
                        "branch_gvcn": b_gvcn
                    })
                    if m_rec.weeks_used:
                        valid_weeks.update([w.strip() for w in m_rec.weeks_used.split(',') if w.strip()])
                        
                valid_weeks = list(valid_weeks)

                if processed_records:
                    top_classes = processed_records[:5] 
                    bottom_classes = processed_records[-5:] if len(processed_records) > 5 else [] 
                    bottom_classes.reverse() 

                    # QUÉT LỖI VI PHẠM
                    violations = db_session.query(WeeklyViolation, ViolationCategory).join(
                        ViolationCategory, WeeklyViolation.violation_id == ViolationCategory.id
                    ).join(
                        WeeklyScore, WeeklyViolation.weekly_score_id == WeeklyScore.id
                    ).filter(WeeklyScore.week.in_(valid_weeks), WeeklyScore.branch_id.in_(branch_map.keys())).all()

                    viol_summary = {}
                    total_violations = 0
                    for v, cat in violations:
                        qty = v.quantity or 1
                        if getattr(cat, 'point_type', 'Điểm trừ') == 'Điểm trừ':
                            viol_summary[cat.name] = viol_summary.get(cat.name, 0) + qty
                            total_violations += qty

                    top_violations_school = sorted(viol_summary.items(), key=lambda x: x[1], reverse=True)[:5] 

                    # =========================================================================
                    # [BẢN VÁ LỖI]: ĐẾM ĐIỂM TỐT TOÀN TRƯỜNG VÀ TOP 5 MÔN HỌC
                    # =========================================================================
                    total_good_points = 0
                    subject_scores = {}

                    # A. Tính Tổng Điểm Tốt Toàn Trường (Từ bảng WeeklyScore để lấy cả điểm nhập tay)
                    weekly_scores = db_session.query(WeeklyScore).filter(
                        WeeklyScore.week.in_(valid_weeks),
                        WeeklyScore.branch_id.in_(branch_map.keys())
                    ).all()

                    for sc in weekly_scores:
                        b_rec = branch_map.get(sc.branch_id)
                        grp = b_rec.group if b_rec else "Nhóm 1"
                        
                        sl_tot = int(sc.count_10 or 0) + int(sc.count_9 or 0)
                        if "2" in str(grp):
                            sl_tot += int(sc.count_8 or 0)
                        total_good_points += sl_tot

                    # B. Tính Top 5 Môn Học (Gắn khóa Y(id) để khớp tuyệt đối với RawScore)
                    safe_valid_weeks = [f"{w}_Y{active_year.id}" for w in valid_weeks]
                    raw_scores = db_session.query(RawScore).filter(RawScore.week.in_(safe_valid_weeks)).all()
                    
                    for rs in raw_scores:
                        subj = rs.subject.strip().title()
                        if not subj or subj.lower() in ["khác", "điểm đã nhập"]: continue
                        
                        b_name = rs.branch_name.strip()
                        c10, c9, c8 = int(rs.c10 or 0), int(rs.c9 or 0), int(rs.c8 or 0)
                        
                        grp = branch_group_map.get(b_name, "")
                        if "1" in str(grp): c8 = 0 
                        
                        points = c10 + c9 + c8
                        if points > 0:
                            subject_scores[subj] = subject_scores.get(subj, 0) + points

                    top_subjects_school = sorted(subject_scores.items(), key=lambda x: x[1], reverse=True)[:5]
                    # =========================================================================

                    analysis_data = {
                        "month_name": selected_month,
                        "weeks_used": ", ".join(valid_weeks),
                        "top_classes": top_classes,
                        "bottom_classes": bottom_classes,
                        "total_violations": total_violations,
                        "top_violations_school": top_violations_school,
                        "total_good_points": total_good_points,
                        "top_subjects_school": top_subjects_school,
                        "total_classes": len(processed_records)
                    }

            return render_template('school_monthly_analysis.html', 
                                   available_months=available_months, 
                                   selected_month=selected_month, 
                                   data=analysis_data)
    except Exception as e:
        import traceback; traceback.print_exc()
        flash(f"Lỗi tải trang báo cáo toàn trường: {e}", "error")
        return redirect(url_for('dashboard'))
# =====================================================================
# MODULE: BÁO CÁO TỔNG QUAN TOÀN TRƯỜNG (HỌC KỲ)
# =====================================================================
@app.route('/school_semester_analysis', methods=['GET'])
def school_semester_analysis():
    try:
        with session_scope() as db_session:
            active_year = db_session.query(SchoolYear).filter_by(is_active=True).first()
            if not active_year:
                flash("Chưa có năm học kích hoạt!", "error")
                return redirect(url_for('dashboard'))

            sems_db = db_session.query(MonthlyRecord.month_name).filter(
                MonthlyRecord.school_year_id == active_year.id,
                MonthlyRecord.month_name.like('Học kỳ%')
            ).distinct().all()
            
            available_semesters = sorted([m[0] for m in sems_db if m[0]])
            selected_semester = request.args.get('semester', available_semesters[-1] if available_semesters else "")
            analysis_data = None

            if selected_semester:
                # 1. Lấy toàn bộ danh sách Chi đoàn
                all_branches = db_session.query(Branch).filter_by(school_year_id=active_year.id).all()
                branch_map = {b.id: b for b in all_branches}
                branch_group_map = {b.name: b.group for b in all_branches}

                # 2. Lấy tất cả các bản ghi điểm học kỳ đã chốt
                records = db_session.query(MonthlyRecord).filter(
                    MonthlyRecord.school_year_id == active_year.id, 
                    MonthlyRecord.month_name == selected_semester
                ).order_by(MonthlyRecord.rank).all()
                
                processed_records = []
                valid_weeks = set()
                
                for s_rec in records:
                    b_rec = branch_map.get(s_rec.branch_id)
                    b_name = str(b_rec.name).strip() if b_rec and b_rec.name else f"Chi đoàn {s_rec.branch_id}"
                    b_gvcn = str(b_rec.gvcn).strip() if b_rec and b_rec.gvcn else "Chưa cập nhật GVCN"

                    processed_records.append({
                        "rank": s_rec.rank,
                        "total_score": round(float(s_rec.total_score or 0), 1),
                        "branch_name": b_name,
                        "branch_gvcn": b_gvcn
                    })
                    
                    # Truy ngược từ Học kỳ -> Các Tháng -> Các Tuần
                    if s_rec.weeks_used:
                        months_used = [m.strip() for m in s_rec.weeks_used.split(',') if m.strip()]
                        for m_name in months_used:
                            m_records = db_session.query(MonthlyRecord).filter(
                                MonthlyRecord.school_year_id == active_year.id,
                                MonthlyRecord.month_name == m_name
                            ).all()
                            for mr in m_records:
                                if mr.weeks_used:
                                    valid_weeks.update([w.strip() for w in mr.weeks_used.split(',') if w.strip()])
                        
                valid_weeks = list(valid_weeks)
                valid_weeks.sort(key=lambda x: int(re.search(r'\d+', str(x)).group()) if re.search(r'\d+', str(x)) else 0)

                if processed_records:
                    top_classes = processed_records[:5] 
                    bottom_classes = processed_records[-5:] if len(processed_records) > 5 else [] 
                    bottom_classes.reverse() 

                    # QUÉT LỖI VI PHẠM (Tính tổng các tuần trong học kỳ)
                    violations = db_session.query(WeeklyViolation, ViolationCategory).join(
                        ViolationCategory, WeeklyViolation.violation_id == ViolationCategory.id
                    ).join(
                        WeeklyScore, WeeklyViolation.weekly_score_id == WeeklyScore.id
                    ).filter(WeeklyScore.week.in_(valid_weeks), WeeklyScore.branch_id.in_(branch_map.keys())).all()

                    viol_summary = {}
                    total_violations = 0
                    for v, cat in violations:
                        qty = v.quantity or 1
                        if getattr(cat, 'point_type', 'Điểm trừ') == 'Điểm trừ':
                            viol_summary[cat.name] = viol_summary.get(cat.name, 0) + qty
                            total_violations += qty

                    top_violations_school = sorted(viol_summary.items(), key=lambda x: x[1], reverse=True)[:5] 

                    # ĐẾM ĐIỂM TỐT TOÀN TRƯỜNG VÀ TOP 5 MÔN HỌC
                    total_good_points = 0
                    subject_scores = {}

                    weekly_scores = db_session.query(WeeklyScore).filter(
                        WeeklyScore.week.in_(valid_weeks),
                        WeeklyScore.branch_id.in_(branch_map.keys())
                    ).all()

                    for sc in weekly_scores:
                        b_rec = branch_map.get(sc.branch_id)
                        grp = b_rec.group if b_rec else "Nhóm 1"
                        
                        sl_tot = int(sc.count_10 or 0) + int(sc.count_9 or 0)
                        if "2" in str(grp):
                            sl_tot += int(sc.count_8 or 0)
                        total_good_points += sl_tot

                    safe_valid_weeks = [f"{w}_Y{active_year.id}" for w in valid_weeks]
                    raw_scores = db_session.query(RawScore).filter(RawScore.week.in_(safe_valid_weeks)).all()
                    
                    for rs in raw_scores:
                        subj = rs.subject.strip().title()
                        if not subj or subj.lower() in ["khác", "điểm đã nhập"]: continue
                        
                        b_name = rs.branch_name.strip()
                        c10, c9, c8 = int(rs.c10 or 0), int(rs.c9 or 0), int(rs.c8 or 0)
                        
                        grp = branch_group_map.get(b_name, "")
                        if "1" in str(grp): c8 = 0 
                        
                        points = c10 + c9 + c8
                        if points > 0:
                            subject_scores[subj] = subject_scores.get(subj, 0) + points

                    top_subjects_school = sorted(subject_scores.items(), key=lambda x: x[1], reverse=True)[:5]

                    analysis_data = {
                        "month_name": selected_semester, # Dùng chung template html nên giữ nguyên key
                        "weeks_used": ", ".join(valid_weeks),
                        "top_classes": top_classes,
                        "bottom_classes": bottom_classes,
                        "total_violations": total_violations,
                        "top_violations_school": top_violations_school,
                        "total_good_points": total_good_points,
                        "top_subjects_school": top_subjects_school,
                        "total_classes": len(processed_records)
                    }

            return render_template('school_monthly_analysis.html', 
                                   available_months=available_semesters, 
                                   selected_month=selected_semester, 
                                   data=analysis_data,
                                   is_semester=True) # Truyền cờ để giao diện biết đây là form Học kỳ
    except Exception as e:
        import traceback; traceback.print_exc()
        flash(f"Lỗi tải trang báo cáo Học kỳ: {e}", "error")
        return redirect(url_for('dashboard'))

# =====================================================================
# MODULE: BÁO CÁO TỔNG QUAN TOÀN TRƯỜNG (CẢ NĂM HỌC)
# =====================================================================
@app.route('/school_yearly_analysis', methods=['GET'])
def school_yearly_analysis():
    try:
        with session_scope() as db_session:
            active_year = db_session.query(SchoolYear).filter_by(is_active=True).first()
            if not active_year:
                flash("Chưa có năm học kích hoạt!", "error")
                return redirect(url_for('dashboard'))

            selected_year_name = f"Năm học {active_year.name}"
            analysis_data = None

            # 1. Lấy toàn bộ danh sách Chi đoàn
            all_branches = db_session.query(Branch).filter_by(school_year_id=active_year.id).all()
            branch_map = {b.id: b for b in all_branches}
            branch_group_map = {b.name: b.group for b in all_branches}

            # 2. Lấy tất cả các bản ghi điểm năm học đã chốt
            records = db_session.query(MonthlyRecord).filter(
                MonthlyRecord.school_year_id == active_year.id, 
                MonthlyRecord.month_name == selected_year_name
            ).order_by(MonthlyRecord.rank).all()
            
            processed_records = []
            valid_weeks = set()
            
            if records:
                for y_rec in records:
                    b_rec = branch_map.get(y_rec.branch_id)
                    b_name = str(b_rec.name).strip() if b_rec and b_rec.name else f"Chi đoàn {y_rec.branch_id}"
                    b_gvcn = str(b_rec.gvcn).strip() if b_rec and b_rec.gvcn else "Chưa cập nhật GVCN"

                    processed_records.append({
                        "rank": y_rec.rank,
                        "total_score": round(float(y_rec.total_score or 0), 1),
                        "branch_name": b_name,
                        "branch_gvcn": b_gvcn
                    })
                    
                    # Truy ngược từ Năm học -> Học kỳ -> Tháng -> Tuần
                    if y_rec.weeks_used:
                        sems_used = [s.strip() for s in y_rec.weeks_used.split(',') if s.strip()]
                        for s_name in sems_used:
                            s_records = db_session.query(MonthlyRecord).filter(
                                MonthlyRecord.school_year_id == active_year.id,
                                MonthlyRecord.month_name == s_name
                            ).all()
                            for sr in s_records:
                                if sr.weeks_used:
                                    months_used = [m.strip() for m in sr.weeks_used.split(',') if m.strip()]
                                    for m_name in months_used:
                                        m_records = db_session.query(MonthlyRecord).filter(
                                            MonthlyRecord.school_year_id == active_year.id,
                                            MonthlyRecord.month_name == m_name
                                        ).all()
                                        for mr in m_records:
                                            if mr.weeks_used:
                                                valid_weeks.update([w.strip() for w in mr.weeks_used.split(',') if w.strip()])
                        
                valid_weeks = list(valid_weeks)
                valid_weeks.sort(key=lambda x: int(re.search(r'\d+', str(x)).group()) if re.search(r'\d+', str(x)) else 0)
                if processed_records:
                    top_classes = processed_records[:5] 
                    bottom_classes = processed_records[-5:] if len(processed_records) > 5 else [] 
                    bottom_classes.reverse() 

                    # QUÉT LỖI VI PHẠM CẢ NĂM
                    violations = db_session.query(WeeklyViolation, ViolationCategory).join(
                        ViolationCategory, WeeklyViolation.violation_id == ViolationCategory.id
                    ).join(
                        WeeklyScore, WeeklyViolation.weekly_score_id == WeeklyScore.id
                    ).filter(WeeklyScore.week.in_(valid_weeks), WeeklyScore.branch_id.in_(branch_map.keys())).all()

                    viol_summary = {}
                    total_violations = 0
                    for v, cat in violations:
                        qty = v.quantity or 1
                        if getattr(cat, 'point_type', 'Điểm trừ') == 'Điểm trừ':
                            viol_summary[cat.name] = viol_summary.get(cat.name, 0) + qty
                            total_violations += qty

                    top_violations_school = sorted(viol_summary.items(), key=lambda x: x[1], reverse=True)[:5] 

                    # ĐẾM ĐIỂM TỐT TOÀN TRƯỜNG VÀ TOP 5 MÔN HỌC
                    total_good_points = 0
                    subject_scores = {}

                    weekly_scores = db_session.query(WeeklyScore).filter(
                        WeeklyScore.week.in_(valid_weeks),
                        WeeklyScore.branch_id.in_(branch_map.keys())
                    ).all()

                    for sc in weekly_scores:
                        b_rec = branch_map.get(sc.branch_id)
                        grp = b_rec.group if b_rec else "Nhóm 1"
                        
                        sl_tot = int(sc.count_10 or 0) + int(sc.count_9 or 0)
                        if "2" in str(grp):
                            sl_tot += int(sc.count_8 or 0)
                        total_good_points += sl_tot

                    safe_valid_weeks = [f"{w}_Y{active_year.id}" for w in valid_weeks]
                    raw_scores = db_session.query(RawScore).filter(RawScore.week.in_(safe_valid_weeks)).all()
                    
                    for rs in raw_scores:
                        subj = rs.subject.strip().title()
                        if not subj or subj.lower() in ["khác", "điểm đã nhập"]: continue
                        
                        b_name = rs.branch_name.strip()
                        c10, c9, c8 = int(rs.c10 or 0), int(rs.c9 or 0), int(rs.c8 or 0)
                        
                        grp = branch_group_map.get(b_name, "")
                        if "1" in str(grp): c8 = 0 
                        
                        points = c10 + c9 + c8
                        if points > 0:
                            subject_scores[subj] = subject_scores.get(subj, 0) + points

                    top_subjects_school = sorted(subject_scores.items(), key=lambda x: x[1], reverse=True)[:5]

                    analysis_data = {
                        "month_name": selected_year_name,
                        "weeks_used": ", ".join(valid_weeks),
                        "top_classes": top_classes,
                        "bottom_classes": bottom_classes,
                        "total_violations": total_violations,
                        "top_violations_school": top_violations_school,
                        "total_good_points": total_good_points,
                        "top_subjects_school": top_subjects_school,
                        "total_classes": len(processed_records)
                    }

            return render_template('school_monthly_analysis.html', 
                                   available_months=[selected_year_name], 
                                   selected_month=selected_year_name, 
                                   data=analysis_data,
                                   is_yearly=True) # Truyền cờ để giao diện biết đây là form Năm học
    except Exception as e:
        import traceback; traceback.print_exc()
        flash(f"Lỗi tải trang tổng kết năm học: {e}", "error")
        return redirect(url_for('dashboard'))    
# =====================================================================
# API: XEM BẢNG XẾP HẠNG TOÀN TRƯỜNG (HỖ TRỢ TUẦN, THÁNG, HỌC KỲ KÈM CHI TIẾT)
# =====================================================================
@app.route('/api/gvcn/leaderboard', methods=['GET'])
@app.route('/api/gvcn/leaderboard/<week_name>', methods=['GET'])
def api_gvcn_leaderboard(week_name=None):
    if not session.get('role'):
        return {"success": False, "error": "Phiên đăng nhập không hợp lệ."}, 401

    try:
        week_name = (week_name or '').strip()

        with session_scope() as db_session:
            active_year = db_session.query(SchoolYear).filter_by(is_active=True).first()
            if not active_year:
                latest_year = db_session.query(SchoolYear).order_by(SchoolYear.id.desc()).first()
                school_year_id = latest_year.id if latest_year else None
            else:
                school_year_id = active_year.id

            if not school_year_id:
                return {"success": False, "error": "Chưa có năm học nào trong hệ thống."}, 200

            # --- TỰ ĐỘNG TÌM TUẦN MỚI NHẤT NẾU KHÔNG CÓ THAM SỐ ---
            if not week_name or week_name in ['undefined', 'null', 'latest']:
                latest_score = db_session.query(WeeklyScore).join(Branch).filter(
                    Branch.school_year_id == school_year_id
                ).order_by(WeeklyScore.id.desc()).first()
                week_name = latest_score.week if latest_score else "Tuần 1"

            branches = db_session.query(Branch).filter_by(school_year_id=school_year_id).all()
            group_1_data = []
            group_2_data = []

            max_tot, max_mon = 14, 4
            try:
                settings = db_session.query(ScoreSettings).filter_by(school_year_id=school_year_id).first()
                if settings:
                    if hasattr(settings, 'max_diem_tot'): max_tot = int(settings.max_diem_tot)
                    if hasattr(settings, 'max_diem_mon'): max_mon = int(settings.max_diem_mon)
            except Exception: pass

            branch_ids = [b.id for b in branches]

            if "Tháng" in week_name:
                month_scores = db_session.query(MonthlyRecord).filter(
                    MonthlyRecord.month_name == week_name, 
                    MonthlyRecord.school_year_id == school_year_id
                ).all()
                weekly_scores_all = db_session.query(WeeklyScore).filter(WeeklyScore.branch_id.in_(branch_ids)).all()

                for branch in branches:
                    m_sc = next((m for m in month_scores if m.branch_id == branch.id), None)
                    grp = str(branch.group or "").strip()
                    tong_diem_tru = 0; tong_diem_cong = 0; tong_diem_tot = 0
                    ghi_chu_gop = []

                    if m_sc and getattr(m_sc, 'weeks_used', None):
                        weeks = [w.strip() for w in m_sc.weeks_used.split(",")]
                        branch_weeks = [w for w in weekly_scores_all if w.branch_id == branch.id and w.week in weeks]
                        for bw in branch_weeks:
                            tong_diem_tru += float(bw.score_tru or bw.score_kem or 0)
                            tong_diem_cong += float(bw.score_cong or 0)
                            if bw.note and bw.note.strip() and bw.note.strip() != 'Không có vi phạm':
                                ghi_chu_gop.append(f"[{bw.week}] {bw.note.strip()}")
                        for w in weeks:
                            tong_diem_tot += calculate_trimmed_good_points_web(db_session, w, branch.name, branch.group, max_mon, max_tot)

                    item = {
                        "branch_name": str(branch.name or ""),
                        "group": grp,
                        "week_rating": f"Tổng hợp {week_name}",
                        "total_score": round(float(m_sc.total_score or 0), 1) if m_sc else 0.0,
                        "note": " | ".join(ghi_chu_gop) if ghi_chu_gop else "Không có vi phạm",
                        "cong": round(tong_diem_cong, 1), "tru": round(tong_diem_tru, 1), "diem_tot": int(tong_diem_tot)
                    }
                    if "1" in grp: group_1_data.append(item)
                    else: group_2_data.append(item)

            elif "Học kỳ" in week_name or "Cả năm" in week_name:
                if "Học kỳ 1" in week_name: target_weeks = [f"Tuần {i}" for i in range(1, 19)]
                elif "Học kỳ 2" in week_name: target_weeks = [f"Tuần {i}" for i in range(19, 38)]
                else: target_weeks = [f"Tuần {i}" for i in range(1, 38)]

                scores = db_session.query(WeeklyScore).filter(WeeklyScore.week.in_(target_weeks), WeeklyScore.branch_id.in_(branch_ids)).all() if target_weeks else []
                for branch in branches:
                    branch_scores = [sc for sc in scores if sc.branch_id == branch.id]
                    tong_diem = sum((sc.total_score or 0) for sc in branch_scores)
                    tong_diem_tru = sum((sc.score_tru or sc.score_kem or 0) for sc in branch_scores)
                    tong_diem_cong = sum((sc.score_cong or 0) for sc in branch_scores)
                    tong_diem_tot = sum(calculate_trimmed_good_points_web(db_session, w, branch.name, branch.group, max_mon, max_tot) for w in target_weeks)
                    ghi_chu_gop = [f"[{sc.week}] {sc.note.strip()}" for sc in branch_scores if sc.note and sc.note.strip() and sc.note.strip() != 'Không có vi phạm']
                    grp = str(branch.group or "").strip()

                    item = {
                        "branch_name": str(branch.name or ""),
                        "group": grp,
                        "week_rating": f"Tổng hợp {week_name}",
                        "total_score": round(float(tong_diem), 1),
                        "note": " | ".join(ghi_chu_gop) if ghi_chu_gop else "Không có vi phạm",
                        "cong": round(tong_diem_cong, 1), "tru": round(tong_diem_tru, 1), "diem_tot": int(tong_diem_tot)
                    }
                    if "1" in grp: group_1_data.append(item)
                    else: group_2_data.append(item)

            else:
                scores = db_session.query(WeeklyScore).filter(WeeklyScore.week == week_name, WeeklyScore.branch_id.in_(branch_ids)).all()
                score_map = {sc.branch_id: sc for sc in scores}
                for branch in branches:
                    sc = score_map.get(branch.id)
                    grp = str(branch.group or "").strip()
                    diem_tot = getattr(sc, 'diem_tot', (sc.count_8 or 0) + (sc.count_9 or 0) + (sc.count_10 or 0)) if sc else 0
                    
                    item = {
                        "branch_name": str(branch.name or ""),
                        "group": grp,
                        "week_rating": str(sc.week_rating or "Bình thường") if sc else "Chưa có dữ liệu",
                        "total_score": round(float(sc.total_score or 0), 1) if sc else 0.0,
                        "note": str(sc.note or "Không có vi phạm") if sc else "Chưa nhập điểm",
                        "cong": float(sc.score_cong or 0) if sc else 0.0,
                        "tru": float(sc.score_tru or sc.score_kem or 0) if sc else 0.0,
                        "diem_tot": int(diem_tot)
                    }
                    if "1" in grp: group_1_data.append(item)
                    else: group_2_data.append(item)

            group_1_data.sort(key=lambda x: x["total_score"], reverse=True)
            group_2_data.sort(key=lambda x: x["total_score"], reverse=True)

            return {"success": True, "week_name": week_name, "group_1": group_1_data, "group_2": group_2_data}, 200

    except Exception as e:
        import traceback; traceback.print_exc()
        return {"success": False, "error": f"{type(e).__name__}: {str(e)}"}, 500

# =====================================================================
# API: TẢI DANH SÁCH THÁNG CHO BỘ LỌC
# =====================================================================
@app.route('/api/gvcn/get_months')
def api_gvcn_get_months():
    try:
        from database.database import session_scope
        from database.models import SchoolYear, MonthlyRecord
        with session_scope() as db_session:
            active_year = db_session.query(SchoolYear).filter_by(is_active=True).first()
            if not active_year: return {"success": False, "months": []}
            
            months_db = db_session.query(MonthlyRecord.month_name).filter(
                MonthlyRecord.school_year_id == active_year.id, MonthlyRecord.month_name.like('Tháng%')
            ).distinct().all()
            
            school_order = ["Tháng 9", "Tháng 10", "Tháng 11", "Tháng 12", "Tháng 1", "Tháng 2", "Tháng 3", "Tháng 4", "Tháng 5"]
            raw_months = [m[0] for m in months_db if m[0]]
            available_months = sorted(raw_months, key=lambda x: school_order.index(x) if x in school_order else 99)
            return {"success": True, "months": available_months}
    except Exception as e:
        return {"success": False, "error": str(e)}
    
# =====================================================================
# API: CẬP NHẬT THÔNG TIN LIÊN HỆ LỚP (ĐÃ VÁ LỖI JSONIFY VÀ BỌC BẢO VỆ KÉP)
# =====================================================================
@app.route('/update_gvcn_info', methods=['POST'])
def update_branch_info():
    if not session.get('role'):
        return {"success": False, "error": "Vui lòng đăng nhập!"}, 401
        
    try:
        data = request.get_json()
        branch_id = data.get('branch_id')
        
        with session_scope() as db_session:
            branch = db_session.query(Branch).filter_by(id=branch_id).first()
            if not branch:
                return {"success": False, "error": "Không tìm thấy Chi đoàn!"}, 404
                
            # Kiểm tra bảo mật nội bộ
            if session.get('role') == 'Giáo viên chủ nhiệm':
                if session.get('username') != branch.name:
                    return {"success": False, "error": "Bạn chỉ có quyền sửa thông tin lớp chủ nhiệm của mình!"}, 403
                    
            # Cập nhật thông tin
            branch.gvcn = data.get('gvcn', '').strip()
            branch.phone_gvcn = data.get('phone_gvcn', '').strip()
            branch.class_monitor = data.get('class_monitor', '').strip()
            branch.phone_monitor = data.get('phone_monitor', '').strip()
            
            # --- [BỔ SUNG VÀ VÁ LỖI]: LƯU SĨ SỐ AN TOÀN ---
            si_so_val = data.get('si_so')
            if si_so_val is not None and str(si_so_val).strip() != "":
                try:
                    branch.si_so = int(si_so_val)
                except ValueError:
                    pass # Bỏ qua nếu dữ liệu đầu vào không phải là số
            # -----------------------------------------------
            
            # Bọc bảo vệ hàm log để chắc chắn không gây sập nếu chưa được định nghĩa
            try:
                log_system_action("CẬP NHẬT HỒ SƠ", f"Cập nhật thông tin liên hệ lớp {branch.name}")
            except Exception:
                pass
                
            return {"success": True, "message": "Đã cập nhật thông tin thành công!"}, 200
            
    except Exception as e:
        import traceback
        traceback.print_exc()
        return {"success": False, "error": f"Lỗi hệ thống: {str(e)}"}, 500
    
@app.route('/bgh/dashboard', methods=['GET'])
def bgh_dashboard():
    if session.get('role') not in ['Quản trị viên', 'Admin', 'Ban Giám hiệu', 'Bí thư Đoàn trường', 'Bí thư']:
        flash("Bạn không có quyền truy cập khu vực điều hành của Ban Giám hiệu!", "error")
        return redirect(url_for('dashboard'))

    # Chỉ lấy tham số week từ URL, bỏ giá trị gán cứng 'Tuần 1'
    selected_week = request.args.get('week')

    with session_scope() as db_session:
        active_year = db_session.query(SchoolYear).filter_by(is_active=True).first()
        school_year_id = active_year.id if active_year else None

        # --- [BẢN VÁ LỖI LÕI]: TÌM CHÍNH XÁC TUẦN MỚI NHẤT BẰNG TOÁN HỌC ---
        if not selected_week:
            import re
            # Lấy tất cả các tuần đang có điểm trong hệ thống
            all_weeks = db_session.query(WeeklyScore.week).join(Branch).filter(
                Branch.school_year_id == school_year_id if school_year_id else True
            ).distinct().all()
            
            if all_weeks:
                # Thuật toán Max: Lọc số và lấy tuần lớn nhất (VD: Tuần 12 sẽ lớn hơn Tuần 2)
                selected_week = max([w[0] for w in all_weeks], key=lambda x: int(re.search(r'\d+', str(x)).group()) if re.search(r'\d+', str(x)) else 0)
            else:
                # Nếu chưa có điểm, lấy theo lịch phân công Sao đỏ
                latest_assign = db_session.query(Assignment).order_by(Assignment.week_number.desc()).first()
                selected_week = f"Tuần {latest_assign.week_number}" if latest_assign else "Tuần 1"
        # ---------------------------------------------------------------------------------

        if not school_year_id:
            return render_template('bgh_dashboard.html', groups_data={}, ai_summary="Chưa kích hoạt năm học.", weeks=[f"Tuần {i}" for i in range(1, 38)])

        branches = db_session.query(Branch).filter_by(school_year_id=school_year_id).all()
        branch_ids = [b.id for b in branches]

        scores = db_session.query(WeeklyScore).filter(
            WeeklyScore.week == selected_week,
            WeeklyScore.branch_id.in_(branch_ids)
        ).all()
        score_map = {sc.branch_id: sc for sc in scores}

        # Gom nhóm dữ liệu theo tên Nhóm (Ví dụ: Nhóm 1, Nhóm 2 hoặc theo Khối)
        groups_dict = {}

        for branch in branches:
            group_name = branch.group if branch.group else "Nhóm Chung"
            if group_name not in groups_dict:
                groups_dict[group_name] = []

            sc = score_map.get(branch.id)
            score = float(sc.total_score or 0) if sc else 0.0
            tru = float(sc.score_tru or sc.score_kem or 0) if sc else 0.0

            groups_dict[group_name].append({
                "branch_name": branch.name,
                "gvcn": branch.gvcn or "Chưa cập nhật",
                "score_tru": tru,
                "total_score": score
            })

        total_school_score = 0
        total_classes_count = 0
        total_violations = 0
        top_classes = []
        warning_classes = []

        # =========================================================================
        # [NÂNG CẤP LÕI]: Sắp xếp tự nhiên (Natural Sort) và Xếp hạng đồng cấp
        # =========================================================================
        
        for group_name in groups_dict:
            import re
            # 1. Sắp xếp: Ưu tiên 1 là Điểm (giảm dần) -> Ưu tiên 2 là Tên lớp (Tự nhiên A-Z: 10A2 trước 10A10)
            groups_dict[group_name].sort(key=lambda x: (
                -x["total_score"], 
                [int(t) if t.isdigit() else t.lower() for t in re.split(r'(\d+)', str(x['branch_name']))]
            ))
            
            # 2. Gán hạng (Rank): Trùng điểm thì được xếp cùng Hạng
            current_rank = 1
            for idx, c in enumerate(groups_dict[group_name]):
                if idx > 0 and c["total_score"] < groups_dict[group_name][idx-1]["total_score"]:
                    current_rank = idx + 1
                c["rank"] = current_rank
                
                total_school_score += c["total_score"]
                total_violations += c["score_tru"]
                total_classes_count += 1

            # 3. Chỉ đưa vào danh sách Lớp dẫn đầu / Cảnh báo nếu các lớp THỰC SỰ đã có điểm (> 0)
            valid_classes = [c for c in groups_dict[group_name] if c['total_score'] > 0]
            if valid_classes:
                top_classes.append({"group": group_name, "class": valid_classes[0]})
                warning_classes.append({"group": group_name, "class": valid_classes[-1]})
        # =========================================================================

        avg_school_score = round(total_school_score / total_classes_count, 1) if total_classes_count > 0 else 0

        # --- [NÂNG CẤP MỚI]: TRỢ LÝ AI PHÂN TÍCH CHUYÊN SÂU & THỰC CHIẾN ---
        all_flattened_classes = []
        for g_name, cls_list in groups_dict.items():
            for c in cls_list:
                c['group_name'] = g_name
                all_flattened_classes.append(c)
        
        # Sắp xếp theo số điểm trừ (score_tru) giảm dần để tìm điểm nóng vi phạm
        all_flattened_classes.sort(key=lambda x: x["score_tru"], reverse=True)
        top_violated_classes = [c for c in all_flattened_classes if c["score_tru"] > 0]

        if total_classes_count == 0 or avg_school_score == 0:
            ai_summary = f"Trong {selected_week}, chưa có dữ liệu điểm số được ghi nhận trên hệ thống để phân tích."
        else:
            insights_sentences = []
            insights_sentences.append(f"📊 **Bức tranh tổng thể {selected_week}:** Điểm trung bình toàn trường đạt **{avg_school_score} điểm** với tổng mức điểm trừ vi phạm ghi nhận là **-{total_violations}đ**.")
            
            if top_violated_classes:
                worst = top_violated_classes[0]
                insights_sentences.append(f"⚠️ **Điểm nóng cần lưu ý:** Lớp **{worst['branch_name']}** ({worst['group_name']}) đang dẫn đầu danh sách vi phạm với mức trừ **-{worst['score_tru']}đ** (GVCN: {worst['gvcn']}).")
            else:
                insights_sentences.append(f"✨ **Tín hiệu tích cực:** Nề nếp toàn trường trong tuần này rất ổn định, không ghi nhận các trường hợp vi phạm nặng.")

            if avg_school_score >= 90:
                insights_sentences.append(f"💡 **Đánh giá của Trợ lý AI:** Nề nếp học sinh duy trì tốt, đề nghị Đoàn trường tiếp tục phát huy và tuyên dương các lớp giữ vững phong độ.")
            else:
                insights_sentences.append(f"💡 **Đề xuất hành động cho BGH:** Đề nghị Đoàn trường phối hợp với Ban giám thị và GVCN các lớp tốp dưới tăng cường nhắc nhở tác phong, đồng phục và giờ giấc trong đầu tuần tới.")

            ai_summary = "<br>".join(insights_sentences)

        return render_template(
            'bgh_dashboard.html',
            selected_week=selected_week,
            avg_school_score=avg_school_score,
            total_violations=total_violations,
            groups_data=groups_dict,
            top_classes=top_classes,
            warning_classes=warning_classes,
            ai_summary=ai_summary,
            weeks=[f"Tuần {i}" for i in range(1, 38)]
        )
# ==========================================
# API: CUNG CẤP DỮ LIỆU ĐIỂM LIVE ĐỂ ĐỒNG BỘ WEB VÀ APP SAO ĐỎ
# ==========================================
@app.route('/api/weekly_scores_json/<week_name>')
def api_weekly_scores_json(week_name):
    # [NÂNG CẤP]: Bổ sung quyền 'Sao đỏ' để App điện thoại có thể lấy được dữ liệu
    if session.get('role') not in ['Quản trị viên', 'Admin', 'Bí thư Đoàn trường', 'Bí thư', 'Sao đỏ', 'Giáo viên chủ nhiệm']:
        return {"success": False, "error": "Unauthorized"}, 401
        
    try:
        with session_scope() as db_session:
            active_year = db_session.query(SchoolYear).filter_by(is_active=True).first()
            if not active_year: return {"success": True, "data": []}
            
            scores = db_session.query(WeeklyScore).join(Branch).filter(
                WeeklyScore.week == week_name,
                Branch.school_year_id == active_year.id
            ).all()
            
            data = []
            for sc in scores:
                data.append({
                    "branch_id": sc.branch_id,
                    "total_score": float(sc.total_score) if sc.total_score is not None else 0.0,
                    "score_tru": float(sc.score_tru) if sc.score_tru is not None else 0.0,
                    "note": sc.note or "",
                    "is_locked": getattr(sc, 'is_locked', False)
                })
            return {"success": True, "data": data}
        
    except Exception as e:
        return {"success": False, "error": str(e)}, 500
    
# ==========================================
# API: XUẤT EXCEL SỔ ĐEN TOÀN TRƯỜNG KÈM BỘ LỌC (ĐÃ VÁ LỖI TRUY VẤN)
# ==========================================
@app.route('/export_filtered_blacklist', methods=['GET'])
def export_filtered_blacklist():
    if session.get('role') not in ['Quản trị viên', 'Admin', 'Bí thư', 'Bí thư Đoàn trường', 'Giáo viên chủ nhiệm', 'Sao đỏ']:
        return redirect(url_for('dashboard'))
        
    try:
        week_name = request.args.get('week', 'Tuần 1').strip()
        keyword = request.args.get('q', '').strip().upper()
        
        with session_scope() as db_session_inner:
            from database.models import WeeklyViolation, WeeklyScore, Branch, ViolationCategory, SchoolYear
            
            active_year_inner = db_session_inner.query(SchoolYear).filter_by(is_active=True).first()
            if not active_year_inner:
                flash("Chưa có năm học nào được kích hoạt!", "error")
                return redirect(url_for('weekly'))
                
            # Truy vấn an toàn, sử dụng điều kiện tương đối với tên tuần để tránh lệch định dạng
            raw_violations = db_session_inner.query(WeeklyViolation, WeeklyScore, Branch, ViolationCategory)\
                .join(WeeklyScore, WeeklyViolation.weekly_score_id == WeeklyScore.id)\
                .join(Branch, WeeklyScore.branch_id == Branch.id)\
                .join(ViolationCategory, WeeklyViolation.violation_id == ViolationCategory.id)\
                .filter(
                    WeeklyViolation.student_name != None,
                    WeeklyViolation.student_name != '',
                    Branch.school_year_id == active_year_inner.id
                ).all()
                
            violation_data = []
            for v, s, b, c in raw_violations:
                # Kiểm tra tuần khớp (hỗ trợ cả trường hợp lệch chữ hoa/thường hoặc khoảng trắng)
                if s.week and s.week.strip().upper() != week_name.upper():
                    continue
                    
                if v.student_name and str(v.student_name).strip() != "":
                    raw_names = str(v.student_name).replace(';', ',').split(',')
                    for raw_n in raw_names:
                        n_clean = raw_n.strip().title()
                        if n_clean:
                            b_name = str(b.name)
                            s_name = str(n_clean)
                            v_name = str(c.name)
                            
                            # Lọc theo từ khóa tìm kiếm trên giao diện (nếu có)
                            if keyword:
                                combined_text = f"{b_name} {s_name} {v_name}".upper()
                                if keyword not in combined_text:
                                    continue
                                    
                            violation_data.append({
                                'branch_name': b_name,
                                'student_name': s_name,
                                'violation_name': v_name,
                                'quantity': int(v.quantity or 1)
                            })
                            
        import openpyxl
        from openpyxl.styles import Font, Alignment, Border, Side
        import io
        from flask import send_file
        
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = f"So_Den_{week_name}"
        
        ws.merge_cells('A1:E1')
        ws['A1'] = "ĐOÀN TRƯỜNG THPT THANH HÒA"
        ws['A1'].font = Font(name="Times New Roman", size=11, bold=True)
        ws.merge_cells('A3:E3')
        ws['A3'] = f"DANH SÁCH SỔ ĐEN KỶ LUẬT - {week_name.upper()}"
        ws['A3'].font = Font(name="Times New Roman", size=14, bold=True)
        ws['A3'].alignment = Alignment(horizontal="center")
        
        if keyword:
            ws.merge_cells('A4:E4')
            ws['A4'] = f"(Đã lọc theo từ khóa: '{keyword}')"
            ws['A4'].font = Font(name="Times New Roman", size=11, italic=True)
            ws['A4'].alignment = Alignment(horizontal="center")
            row_offset = 6
        else:
            row_offset = 5
            
        headers = ["STT", "Chi đoàn", "Họ và Tên Học Sinh", "Lỗi Vi Phạm", "Số Lần"]
        thin = Side(border_style="thin", color="000000")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        
        for col, h in enumerate(headers, 1):
            c = ws.cell(row=row_offset, column=col, value=h)
            c.font = Font(name="Times New Roman", size=11, bold=True)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = border
            
        for idx, item in enumerate(violation_data, 1):
            row_idx = idx + row_offset
            c1 = ws.cell(row=row_idx, column=1, value=idx)
            c2 = ws.cell(row=row_idx, column=2, value=item['branch_name'])
            c3 = ws.cell(row=row_idx, column=3, value=item['student_name'])
            c4 = ws.cell(row=row_idx, column=4, value=item['violation_name'])
            c5 = ws.cell(row=row_idx, column=5, value=item['quantity'])
            
            for cell in [c1, c2, c3, c4, c5]:
                cell.font = Font(name="Times New Roman", size=11)
                cell.border = border
            c1.alignment = Alignment(horizontal="center")
            c2.alignment = Alignment(horizontal="center")
            c5.alignment = Alignment(horizontal="center")
            
        ws.column_dimensions['A'].width = 6
        ws.column_dimensions['B'].width = 15
        ws.column_dimensions['C'].width = 25
        ws.column_dimensions['D'].width = 35
        ws.column_dimensions['E'].width = 10
        
        log_system_action("XUẤT EXCEL", f"Xuất Sổ đen {week_name} theo bản lọc")
        out = io.BytesIO()
        wb.save(out)
        out.seek(0)
        return send_file(out, download_name=f"So_Den_{week_name.replace(' ', '_')}.xlsx", as_attachment=True)
        
    except Exception as e:
        import traceback; traceback.print_exc()
        flash(f"Lỗi xuất Excel: {e}", "error")
        return redirect(url_for('weekly'))
    
# ==========================================
# API: XÓA NHẬT KÝ HỆ THỐNG ĐỂ GIẢI PHÓNG TÀI NGUYÊN
# ==========================================
@app.route('/clear_action_logs', methods=['POST'])
def clear_action_logs():
    # Kiểm tra quyền bảo mật: Chỉ Quản trị viên, Admin hoặc Bí thư Đoàn trường mới được xóa
    if session.get('role') not in ['Quản trị viên', 'Admin', 'Bí thư', 'Bí thư Đoàn trường']:
        flash("Bạn không có quyền thực hiện thao tác này!", "error")
        return redirect(url_for('dashboard'))
    
    try:
        with session_scope() as db_session:
            # Xóa toàn bộ dữ liệu trong bảng ActionLog
            db_session.query(ActionLog).delete()
            db_session.commit()
            
        # Ghi nhận log mới cho thao tác dọn dẹp hệ thống vừa thực hiện
        log_system_action("DỌN DẸP HỆ THỐNG", f"Tài khoản {session.get('username')} đã dọn dẹp sạch toàn bộ lịch sử nhật ký.")
        
        flash("✅ Đã dọn dẹp sạch nhật ký hệ thống để giải phóng tài nguyên!", "success")
    except Exception as e:
        import traceback
        traceback.print_exc()
        flash(f"Lỗi khi xóa nhật ký: {str(e)}", "error")
        
    return redirect(url_for('action_logs'))
# ==========================================
# MODULE: GVCN TRA CỨU & XUẤT SỔ ĐEN CỦA LỚP
# ==========================================
@app.route('/api/class_blacklist')
def api_class_blacklist():
    if session.get('role') not in ['Giáo viên chủ nhiệm', 'Quản trị viên', 'Admin', 'Ban Giám hiệu', 'Bí thư Đoàn trường', 'Bí thư']:
        return {"success": False, "error": "Không có quyền truy cập"}
        
    branch_id = request.args.get('branch_id', type=int)
    time_mode = request.args.get('time_mode', 'week') 
    time_value = request.args.get('time_value', '')

    try:
        with session_scope() as db_session:
            query = db_session.query(
                WeeklyViolation, WeeklyScore, ViolationCategory
            ).join(WeeklyScore, WeeklyViolation.weekly_score_id == WeeklyScore.id)\
             .join(ViolationCategory, WeeklyViolation.violation_id == ViolationCategory.id)\
             .filter(
                WeeklyScore.branch_id == branch_id,
                WeeklyViolation.student_name != None,
                WeeklyViolation.student_name != ''
             )

            if time_mode == 'week' and time_value:
                query = query.filter(WeeklyScore.week == time_value)
            elif time_mode == 'month' and time_value:
                active_year = db_session.query(SchoolYear).filter_by(is_active=True).first()
                if active_year:
                    m_rec_general = db_session.query(MonthlyRecord).filter_by(
                        school_year_id=active_year.id, month_name=time_value
                    ).first()
                    
                    if m_rec_general and m_rec_general.weeks_used:
                        valid_weeks = [w.strip() for w in m_rec_general.weeks_used.split(',') if w.strip()]
                        query = query.filter(WeeklyScore.week.in_(valid_weeks))
                    else:
                        query = query.filter(WeeklyScore.week == 'NONE')

            results = query.order_by(WeeklyScore.id.desc()).all()
            
            # --- ĐÃ SỬA TÊN BIẾN THÀNH 'data' CHO ĐỒNG BỘ ---
            data = [] 
            for v, sc, c in results: 
                if v.student_name and str(v.student_name).strip() != "":
                    raw_names = str(v.student_name).replace(';', ',').split(',')
                    valid_names = [n.strip().title() for n in raw_names if n.strip()]
                    
                    # [THUẬT TOÁN CHIA ĐỀU LỖI VÀ ĐIỂM TRỪ CHO GVCN]
                    num_names = len(valid_names)
                    qty_per_student = max(1, v.quantity // num_names) if num_names > 0 else v.quantity
                    
                    for n_clean in valid_names:
                        data.append({  # Gọi đúng biến data.append
                            'week': sc.week,
                            'student_name': n_clean,
                            'violation_name': c.name,
                            'quantity': qty_per_student, 
                            'penalty': float(c.penalty_points * qty_per_student) if getattr(c, 'point_type', 'Điểm trừ') != 'Điểm cộng' else 0
                        })
            return {"success": True, "data": data}
    except Exception as e:
        import traceback; traceback.print_exc()
        return {"success": False, "error": str(e)}

@app.route('/export_class_blacklist')
def export_class_blacklist():
    if session.get('role') not in ['Giáo viên chủ nhiệm', 'Quản trị viên', 'Admin', 'Ban Giám hiệu', 'Bí thư Đoàn trường', 'Bí thư']:
        flash("Bạn không có quyền!", "error")
        return redirect(url_for('class_dashboard'))
        
    branch_id = request.args.get('branch_id', type=int)
    time_mode = request.args.get('time_mode', 'week')
    time_value = request.args.get('time_value', '')
    
    try:
        with session_scope() as db_session:
            branch = db_session.query(Branch).filter_by(id=branch_id).first()
            if not branch:
                flash("Không tìm thấy lớp!", "error")
                return redirect(url_for('class_dashboard'))
                
            query = db_session.query(
                WeeklyViolation, WeeklyScore, ViolationCategory
            ).join(WeeklyScore, WeeklyViolation.weekly_score_id == WeeklyScore.id)\
             .join(ViolationCategory, WeeklyViolation.violation_id == ViolationCategory.id)\
             .filter(
                WeeklyScore.branch_id == branch_id,
                WeeklyViolation.student_name != None,
                WeeklyViolation.student_name != ''
             )

            if time_mode == 'week' and time_value:
                query = query.filter(WeeklyScore.week == time_value)
            elif time_mode == 'month' and time_value:
                active_year = db_session.query(SchoolYear).filter_by(is_active=True).first()
                if active_year:
                    m_rec = db_session.query(MonthlyRecord).filter_by(
                        school_year_id=active_year.id, month_name=time_value
                    ).first()
                    if m_rec and m_rec.weeks_used:
                        valid_weeks = [w.strip() for w in m_rec.weeks_used.split(',') if w.strip()]
                        query = query.filter(WeeklyScore.week.in_(valid_weeks))
                    else:
                        query = query.filter(WeeklyScore.week == 'NONE')

            results = query.order_by(WeeklyScore.id.desc()).all()
            
            violation_data = []
            for v, sc, c in results: 
                if v.student_name and str(v.student_name).strip() != "":
                    raw_names = str(v.student_name).replace(';', ',').split(',')
                    valid_names = [n.strip().title() for n in raw_names if n.strip()]
                    
                    # [THUẬT TOÁN CHIA ĐỀU CHO FILE EXCEL GVCN TẢI VỀ]
                    num_names = len(valid_names)
                    qty_per_student = max(1, v.quantity // num_names) if num_names > 0 else v.quantity
                    
                    for n_clean in valid_names:
                        violation_data.append({ 
                            'week': sc.week,
                            'student_name': n_clean,
                            'violation_name': c.name,
                            'quantity': qty_per_student,
                            'penalty': float(c.penalty_points * qty_per_student) if getattr(c, 'point_type', 'Điểm trừ') != 'Điểm cộng' else 0
                        })
                        
            # Tạo Excel
            import openpyxl
            from openpyxl.styles import Font, Alignment, Border, Side
            import io
            from flask import send_file
            
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "So_Den_Cua_Lop"
            
            ws.merge_cells('A1:F1')
            ws['A1'] = "ĐOÀN TRƯỜNG THPT THANH HÒA"
            ws['A1'].font = Font(name="Times New Roman", size=11, bold=True)
            
            ws.merge_cells('A3:F3')
            ws['A3'] = f"DANH SÁCH HỌC SINH VI PHẠM KỶ LUẬT - LỚP {branch.name}"
            ws['A3'].font = Font(name="Times New Roman", size=14, bold=True)
            ws['A3'].alignment = Alignment(horizontal="center")
            
            ws.merge_cells('A4:F4')
            ws['A4'] = f"Thời gian thống kê: {time_value}"
            ws['A4'].font = Font(name="Times New Roman", size=12, italic=True)
            ws['A4'].alignment = Alignment(horizontal="center")
            
            # --- ĐÃ BỔ SUNG CỘT ĐIỂM TRỪ CHO ĐỒNG BỘ VỚI WEB ---
            headers = ["STT", "Thời gian", "Họ và Tên", "Lỗi Vi Phạm", "Số Lần", "Điểm Trừ"]
            thin = Side(border_style="thin", color="000000")
            border = Border(left=thin, right=thin, top=thin, bottom=thin)
            
            for col, h in enumerate(headers, 1):
                c = ws.cell(row=6, column=col, value=h)
                c.font = Font(name="Times New Roman", size=12, bold=True)
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.border = border
                
            for idx, item in enumerate(violation_data, 1):
                row_idx = idx + 6
                c1 = ws.cell(row=row_idx, column=1, value=idx)
                c2 = ws.cell(row=row_idx, column=2, value=item['week'])
                c3 = ws.cell(row=row_idx, column=3, value=item['student_name'])
                c4 = ws.cell(row=row_idx, column=4, value=item['violation_name'])
                c5 = ws.cell(row=row_idx, column=5, value=item['quantity'])
                c6 = ws.cell(row=row_idx, column=6, value=f"-{item['penalty']}đ")
                
                for cell in [c1, c2, c3, c4, c5, c6]:
                    cell.font = Font(name="Times New Roman", size=12)
                    cell.border = border
                c1.alignment = Alignment(horizontal="center")
                c2.alignment = Alignment(horizontal="center")
                c5.alignment = Alignment(horizontal="center")
                c6.alignment = Alignment(horizontal="center", wrap_text=True)
                
            ws.column_dimensions['A'].width = 6
            ws.column_dimensions['B'].width = 15
            ws.column_dimensions['C'].width = 25
            ws.column_dimensions['D'].width = 35
            ws.column_dimensions['E'].width = 10
            ws.column_dimensions['F'].width = 10
            
            out = io.BytesIO()
            wb.save(out)
            out.seek(0)
            
            filename = f"So_Den_{branch.name}_{time_value}.xlsx".replace(" ", "_")
            return send_file(out, download_name=filename, as_attachment=True)
            
    except Exception as e:
        import traceback; traceback.print_exc()
        flash(f"Lỗi xuất Excel: {str(e)}", "error")
        return redirect(url_for('class_dashboard'))
# ==========================================
# MODULE: TRUNG TÂM QUẢN LÝ HỒ SƠ PHÚC KHẢO TOÀN TRƯỜNG
# ==========================================
@app.route('/appeals', methods=['GET'])
def manage_appeals():
    # Chỉ cho phép Admin, BGH và Bí thư truy cập
    if session.get('role') not in ['Quản trị viên', 'Admin', 'Ban Giám hiệu', 'Bí thư Đoàn trường', 'Bí thư']:
        flash("Bạn không có quyền truy cập trang quản lý phúc khảo!", "error")
        return redirect(url_for('dashboard'))

    try:
        with session_scope() as db_session:
            active_year = db_session.query(SchoolYear).filter_by(is_active=True).first()
            if not active_year:
                flash("Chưa có năm học kích hoạt!", "error")
                return redirect(url_for('dashboard'))

            # Lấy dữ liệu cho bộ lọc
            branches = db_session.query(Branch).filter_by(school_year_id=active_year.id).all()
            
            weeks_db = db_session.query(WeeklyScore.week).join(Branch).filter(Branch.school_year_id == active_year.id).distinct().all()
            available_weeks = sorted([w[0] for w in weeks_db], key=lambda x: int(re.search(r'\d+', x).group()) if re.search(r'\d+', x) else 0)

            # Lấy tham số tìm kiếm từ giao diện
            search_name = request.args.get('search_name', '').strip().lower()
            search_branch = request.args.get('search_branch', '')
            search_week = request.args.get('search_week', '')
            search_status = request.args.get('search_status', 'all') # all, pending, resolved

            # Truy vấn cơ sở dữ liệu các Tuần CÓ đánh dấu phúc khảo
            query = db_session.query(WeeklyScore).join(Branch).filter(
                Branch.school_year_id == active_year.id,
                WeeklyScore.is_appealed == True
            )

            # Áp dụng bộ lọc cơ sở
            if search_branch and search_branch.isdigit():
                query = query.filter(Branch.id == int(search_branch))
            if search_week:
                query = query.filter(WeeklyScore.week == search_week)
            if search_name:
                query = query.filter(WeeklyScore.appeal_reason.ilike(f"%{search_name}%"))

            if search_status == 'pending':
                query = query.filter((WeeklyScore.appeal_response == None) | (WeeklyScore.appeal_response == ""))
            elif search_status == 'resolved':
                query = query.filter(WeeklyScore.appeal_response != None, WeeklyScore.appeal_response != "")

            appealed_scores = query.order_by(WeeklyScore.id.desc()).all()

            # Bóc tách chuỗi phúc khảo thành các bản ghi chi tiết
            appeal_records = []
            
            # Khởi tạo data Ngân hàng lỗi để Javascript dùng tính điểm hoàn tự động
            violation_bank = db_session.query(ViolationCategory).filter_by(school_year_id=active_year.id).all()

            for sc in appealed_scores:
                if not sc.appeal_reason: continue
                
                pattern = r'\[\d{2}/\d{2}/\d{4} \d{2}:\d{2}\]'
                timestamps = re.findall(pattern, sc.appeal_reason)
                segments = re.split(pattern, sc.appeal_reason)[1:] 
                
                # Chạy ngược để hiển thị khiếu nại mới nhất lên đầu
                for i in range(len(timestamps)-1, -1, -1):
                    time_str = timestamps[i].strip('[]')
                    content = segments[i].strip().strip('|').strip()
                    
                    errors_part = ""
                    reason_part = content
                    
                    match = re.search(r'Phúc khảo các lỗi:\s*(.*?)\s*\|\s*Lý do:(.*)', content, re.IGNORECASE)
                    if match:
                        errors_part = match.group(1).strip()
                        reason_part = match.group(2).strip()

                    # Lọc lại Tên học sinh một lần nữa trên chuỗi đã bóc tách cho chính xác
                    if search_name and search_name not in errors_part.lower() and search_name not in reason_part.lower():
                        continue
                    
                    status_text = "Đang chờ xử lý"
                    badge_class = "warning text-dark"
                    if sc.appeal_response:
                        if "ĐÃ DUYỆT" in sc.appeal_response:
                            status_text = "Đã duyệt"
                            badge_class = "success"
                        elif "TỪ CHỐI" in sc.appeal_response:
                            status_text = "Từ chối"
                            badge_class = "danger"
                    
                    appeal_records.append({
                        'score_id': sc.id,
                        'branch_name': sc.branch.name,
                        'week': sc.week,
                        'time': time_str,
                        'errors_raw': errors_part,
                        'reason': reason_part,
                        'status': status_text,
                        'badge': badge_class,
                        'response': sc.appeal_response if sc.appeal_response else "",
                        'raw_reason': sc.appeal_reason # Dùng để nhét vào form Xử lý
                    })

            return render_template('appeals.html', 
                                   branches=branches,
                                   available_weeks=available_weeks,
                                   appeal_records=appeal_records,
                                   search_name=search_name,
                                   search_branch=search_branch,
                                   search_week=search_week,
                                   search_status=search_status,
                                   violation_bank=violation_bank,
                                   active_year=active_year)
    except Exception as e:
        import traceback; traceback.print_exc()
        flash(f"Lỗi tải danh sách phúc khảo: {e}", "error")
        return redirect(url_for('dashboard'))

# ========================================================
# [TÍNH NĂNG MỚI]: BẢNG ĐIỆN TỬ GVCN - CẬP NHẬT TỪ QUẢN TRỊ
# ========================================================
@app.context_processor
def inject_slogan():
    import os
    slogan_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config", "slogan.txt")
    # Câu chào mặc định nếu file chưa được tạo
    slogan_text = "CHÀO MỪNG NĂM HỌC MỚI - ĐOÀN VIÊN THANH NIÊN TRƯỜNG THPT THANH HÒA TIÊN PHONG, BẢN LĨNH, SÁNG TẠO!"
    if os.path.exists(slogan_file):
        with open(slogan_file, 'r', encoding='utf-8') as f:
            content = f.read().strip()
            if content:
                slogan_text = content
    return dict(global_slogan_text=slogan_text)

@app.route('/update_slogan', methods=['POST'])
def update_slogan():
    if session.get('role') not in ['Quản trị viên', 'Admin', 'Bí thư Đoàn trường']:
        flash("Bạn không có quyền thay đổi thông báo!", "error")
        return redirect(request.referrer)
        
    new_text = request.form.get('slogan_text', '').strip()
    import os
    config_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config")
    os.makedirs(config_dir, exist_ok=True) # Tự động tạo thư mục config nếu chưa có
    slogan_file = os.path.join(config_dir, "slogan.txt")
    
    try:
        with open(slogan_file, 'w', encoding='utf-8') as f:
            f.write(new_text)
        flash("Đã phát sóng nội dung mới lên toàn bộ App GVCN thành công!", "success")
    except Exception as e:
        flash(f"Lỗi hệ thống khi lưu: {e}", "error")
        
    return redirect(request.referrer)
@app.route('/reset_system_data', methods=['POST'])
def reset_system_data():
    if session.get('role') not in ['Quản trị viên', 'Admin', 'Bí thư', 'Bí thư Đoàn trường']:
        flash("⛔ Bạn không có quyền thực hiện thao tác này!", "error")
        return redirect(url_for('school_years'))

    admin_password = request.form.get('admin_password', '').strip()
    current_username = session.get('username')
    
    print(f"--- ĐANG THỰC HIỆN RESET CHO USER: {current_username} ---")

    try:
        with session_scope() as db_session:
            current_user = db_session.query(User).filter_by(username=current_username).first()
            
            if not current_user:
                flash("❌ Không tìm thấy thông tin tài khoản hiện tại trong CSDL!", "error")
                return redirect(url_for('school_years'))

            # Kiểm tra mật khẩu (Hỗ trợ cả trường hợp khớp trực tiếp hoặc gõ cứng "1" để test nhanh)
            if current_user.password_hash != admin_password and admin_password != "1":
                print(f"❌ Sai mật khẩu! Mật khẩu nhập: [{admin_password}], Mật khẩu trong DB: [{current_user.password_hash}]")
                flash("❌ Mật khẩu quản trị không chính xác! Thao tác reset bị hủy bỏ.", "error")
                return redirect(url_for('school_years'))

            print("✅ Xác thực mật khẩu thành công. Đang tiến hành xóa dữ liệu...")

            # Thực hiện xóa dữ liệu các bảng thi đua
            db_session.query(WeeklyViolation).delete()
            db_session.query(StarEvaluation).delete()
            db_session.query(WeeklyScore).delete()
            db_session.query(MonthlyRecord).delete()
            db_session.query(Assignment).delete()
            db_session.query(ActionLog).delete()
            db_session.query(GVCNAttendance).delete()
            db_session.query(RawScore).delete()
            
            db_session.commit()
            print("✅ Đã commit xóa dữ liệu thành công!")

            log_system_action("RESET HỆ THỐNG", f"Tài khoản {current_username} đã reset toàn bộ dữ liệu thi đua.")
            flash("🔄 ĐÃ RESET HỆ THỐNG THÀNH CÔNG! Toàn bộ dữ liệu điểm số đã được làm sạch.", "success")
            
    except Exception as e:
        import traceback
        traceback.print_exc() # In chi tiết lỗi ra màn hình Terminal
        flash(f"❌ Lỗi ngoại lệ khi reset: {str(e)}", "error")
        
    return redirect(url_for('school_years'))

# =====================================================================
# MODULE: HỒ SƠ CHỦ NHIỆM (XEM TRƯỚC VÀ XUẤT EXCEL) - Dành riêng Admin
# =====================================================================
@app.route('/homeroom_portfolio', methods=['GET', 'POST'])
def homeroom_portfolio():
    if session.get('role') not in ['Quản trị viên', 'Admin', 'Bí thư Đoàn trường', 'Bí thư']:
        flash("⛔ Từ chối truy cập: Chỉ Quản trị viên mới được sử dụng chức năng này!", "error")
        return redirect(url_for('dashboard'))

    try:
        with session_scope() as db_session:
            active_year = db_session.query(SchoolYear).filter_by(is_active=True).first()
            if not active_year:
                flash("Chưa có năm học kích hoạt!", "error")
                return redirect(url_for('dashboard'))

            branches = db_session.query(Branch).filter_by(school_year_id=active_year.id).order_by(Branch.name).all()
            
            # Lấy danh sách Tháng
            months_db = db_session.query(MonthlyRecord.month_name).filter(
                MonthlyRecord.school_year_id == active_year.id, MonthlyRecord.month_name.like('Tháng%')
            ).distinct().all()
            school_order = ["Tháng 9", "Tháng 10", "Tháng 11", "Tháng 12", "Tháng 1", "Tháng 2", "Tháng 3", "Tháng 4", "Tháng 5"]
            available_months = sorted([m[0] for m in months_db if m[0]], key=lambda x: school_order.index(x) if x in school_order else 99)

            # Lấy danh sách Tuần
            weeks_db = db_session.query(WeeklyScore.week).join(Branch).filter(Branch.school_year_id == active_year.id).distinct().all()
            available_weeks = sorted([w[0] for w in weeks_db if w[0]], key=lambda x: int(re.search(r'\d+', x).group()) if re.search(r'\d+', x) else 0)

            # Tham số bộ lọc
            selected_branch_id = request.form.get('branch_id') or request.args.get('branch_id')
            selected_time = request.form.get('time_filter') or request.args.get('time_filter')
            
            if not selected_branch_id and branches: selected_branch_id = branches[0].id
            if not selected_time: 
                selected_time = available_months[-1] if available_months else (available_weeks[-1] if available_weeks else "")
            
            selected_branch_id = int(selected_branch_id) if selected_branch_id else 0
            selected_branch = db_session.query(Branch).filter_by(id=selected_branch_id).first()

            matrix_data = []
            target_weeks = []
            
            if selected_branch and selected_time:
                # Nếu người dùng chọn THÁNG -> Bung ra 4-5 tuần
                if selected_time.startswith("Tháng"):
                    month_rec = db_session.query(MonthlyRecord).filter_by(
                        branch_id=selected_branch.id, month_name=selected_time, school_year_id=active_year.id
                    ).first()
                    
                    if month_rec and month_rec.weeks_used:
                        raw_weeks = [w.strip() for w in month_rec.weeks_used.split(',') if w.strip()]
                        target_weeks = sorted(raw_weeks, key=lambda x: int(re.search(r'\d+', x).group()) if re.search(r'\d+', x) else 0)
                    target_weeks = target_weeks[:5]
                # Nếu người dùng chọn TUẦN -> Chỉ xuất 1 tuần đó
                else:
                    target_weeks = [selected_time]
                
                row_labels = [
                    ("Số học sinh đi muộn", ['muộn', 'trễ']),
                    ("Số học sinh bỏ tiết", ['bỏ tiết', 'trốn', 'vắng', 'nghỉ']),
                    ("Số không chuẩn bị bài", ['không học', 'không chuẩn bị', 'không thuộc']),
                    ("Số bị dưới 5,0 hoặc nhận xét loại: yếu, kém", ['điểm kém', 'yếu', '0 điểm']),
                    ("Mắc thái độ sai", ['thái độ', 'vô lễ', 'ồn', 'nói chuyện', 'đùa giỡn']),
                    ("Số điểm tốt", []), 
                    ("Số việc tốt", ['việc tốt', 'nhặt được']),
                    ("HS được khen", ['khen', 'tuyên dương']),
                    ("HS bị phê bình", ['phê bình', 'khiển trách']),
                    ("Số tiết trống", ['trống', 'giáo viên vắng']),
                    ("Số tiết tự quản tốt", ['tự quản']),
                    ("Xếp loại cả lớp", []) 
                ]

                matrix_dict = {label: [] for label, _ in row_labels}

                for week in target_weeks:
                    sc = db_session.query(WeeklyScore).filter_by(branch_id=selected_branch.id, week=week).first()
                    if not sc:
                        for key in matrix_dict: matrix_dict[key].append("")
                        continue

                    counts = {label: 0 for label, _ in row_labels}
                    
                    vios = db_session.query(WeeklyViolation, ViolationCategory).join(
                        ViolationCategory, WeeklyViolation.violation_id == ViolationCategory.id
                    ).filter(WeeklyViolation.weekly_score_id == sc.id).all()

                    for v, cat in vios:
                        c_name = cat.name.lower()
                        qty = v.quantity or 1
                        for label, keywords in row_labels:
                            if keywords and any(k in c_name for k in keywords):
                                counts[label] += qty
                                break 

                    counts["Số điểm tốt"] = int(sc.count_8 or 0) + int(sc.count_9 or 0) + int(sc.count_10 or 0)
                    counts["Xếp loại cả lớp"] = sc.week_rating or "Bình thường"

                    for label, _ in row_labels:
                        val = counts[label]
                        if isinstance(val, int) and val == 0: val = "" 
                        matrix_dict[label].append(val)

                for label, _ in row_labels:
                    vals = matrix_dict[label] + [""] * (5 - len(target_weeks))
                    matrix_data.append({"label": label, "values": vals})

            return render_template('homeroom_portfolio.html', 
                                   branches=branches, 
                                   available_months=available_months,
                                   available_weeks=available_weeks,
                                   selected_branch=selected_branch,
                                   selected_time=selected_time,
                                   target_weeks=target_weeks,
                                   matrix_data=matrix_data)
    except Exception as e:
        import traceback; traceback.print_exc()
        flash(f"Lỗi: {e}", "error")
        return redirect(url_for('dashboard'))

@app.route('/export_homeroom_portfolio', methods=['POST'])
def export_homeroom_portfolio():
    if session.get('role') not in ['Quản trị viên', 'Admin', 'Bí thư Đoàn trường', 'Bí thư']:
        return redirect(url_for('dashboard'))

    try:
        matrix_data_json = request.form.get('matrix_data')
        branch_name = request.form.get('branch_name', 'Lop_Khong_Ten')
        time_name = request.form.get('time_name', 'Thoi_gian')
        target_weeks = json.loads(request.form.get('target_weeks', '[]'))
        data = json.loads(matrix_data_json) if matrix_data_json else []

        if not data:
            flash("Không có dữ liệu để xuất!", "error")
            return redirect(url_for('homeroom_portfolio'))

        wb = openpyxl.Workbook(); ws = wb.active; ws.title = "So_Ket_Tuan"
        
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_margins = PageMargins(left=0.5, right=0.5, top=0.5, bottom=0.5)

        font_title = Font(name='Times New Roman', size=16, bold=True)
        font_header = Font(name='Times New Roman', size=12, bold=True)
        font_normal = Font(name='Times New Roman', size=12)
        align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        ws.merge_cells('A1:F1'); ws['A1'] = "ĐOÀN TRƯỜNG THPT THANH HÒA"; ws['A1'].font = Font(name='Times New Roman', size=12, bold=True); ws['A1'].alignment = align_left
        ws.merge_cells('A3:F3'); ws['A3'] = f"SƠ KẾT HÀNG TUẦN - LỚP {branch_name.upper()}"; ws['A3'].font = font_title; ws['A3'].alignment = align_center
        ws.merge_cells('A4:F4'); ws['A4'] = f"Kỳ đánh giá: {time_name}"; ws['A4'].font = Font(name='Times New Roman', size=12, italic=True); ws['A4'].alignment = align_center

        start_row = 6
        ws.cell(row=start_row, column=1, value="Nội dung đánh giá").font = font_header; ws.cell(row=start_row, column=1).alignment = align_center; ws.cell(row=start_row, column=1).border = thin_border
        
        for i in range(5): 
            col = i + 2
            val = target_weeks[i] if i < len(target_weeks) else ""
            cell = ws.cell(row=start_row, column=col, value=val)
            cell.font = font_header; cell.alignment = align_center; cell.border = thin_border

        current_row = start_row + 1
        for item in data:
            c_label = ws.cell(row=current_row, column=1, value=item['label'])
            c_label.font = font_normal; c_label.alignment = align_left; c_label.border = thin_border
            
            for i, val in enumerate(item['values']):
                c_val = ws.cell(row=current_row, column=i+2, value=val)
                c_val.font = font_normal; c_val.alignment = align_center; c_val.border = thin_border
                
            ws.row_dimensions[current_row].height = 25 
            current_row += 1

        ws.column_dimensions['A'].width = 38
        for col_letter in ['B', 'C', 'D', 'E', 'F']:
            ws.column_dimensions[col_letter].width = 18

        current_row += 2
        ws.merge_cells(start_row=current_row, start_column=5, end_row=current_row, end_column=6)
        ws.cell(row=current_row, column=5, value="Giáo viên Chủ nhiệm").font = font_header; ws.cell(row=current_row, column=5).alignment = align_center

        log_system_action("XUẤT EXCEL", f"Xuất Hồ sơ chủ nhiệm lớp {branch_name} - {time_name}")
        out = io.BytesIO(); wb.save(out); out.seek(0)
        return send_file(out, download_name=f"Ho_So_{branch_name}_{time_name.replace(' ', '_')}.xlsx", as_attachment=True)

    except Exception as e:
        flash(f"Lỗi xuất Excel: {e}", "error")
        return redirect(url_for('homeroom_portfolio'))
    
if __name__ == "__main__":
    auto_init_accounts()
    init_db()
    create_mock_admin()
    
    # Khởi động luồng chạy ngầm sao lưu dữ liệu tự động
    import threading
    threading.Thread(target=background_auto_backup, daemon=True).start()
    
    print("=========================================================")
    print("🌟 BẢN CẬP NHẬT HOÀN HẢO - SẴN SÀNG THỰC CHIẾN 🌟")
    print("=========================================================")
    print("🚀 Máy chủ Web (Phiên bản chịu tải cao) đang khởi động...")
    print("🌍 Sẵn sàng đón nhận 100+ kết nối cùng lúc.")
    print("👉 Hãy mở Cloudflare Tunnel hoặc truy cập Local IP tại cổng: 8080")
    
    # [TINH CHỈNH NHỎ]: Tự động nhận diện Port từ Server hoặc mặc định là 8080
    import os
    port = int(os.environ.get("PORT", 8080))
    
    # =========================================================
    # [NÂNG CẤP LÕI]: SỬ DỤNG WAITRESS THAY VÌ APP.RUN ĐỂ CHỊU TẢI
    # =========================================================
    try:
        from waitress import serve
        # Mở 100 luồng (threads) để xử lý song song 100 thao tác cùng 1 mili-giây
        serve(app, host='0.0.0.0', port=port, threads=100)
    except ImportError:
        print("⚠️ CẢNH BÁO: Chưa cài đặt thư viện Waitress!")
        print("Đang chạy tạm bằng máy chủ thử nghiệm (Chậm hơn). Hãy gõ lệnh: pip install waitress")
        app.run(debug=False, use_reloader=False, host='0.0.0.0', port=port)