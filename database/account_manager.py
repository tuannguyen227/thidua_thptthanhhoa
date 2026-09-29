import os
import json

ACCOUNTS_FILE = os.path.join(os.getcwd(), "data", "accounts.json")

def load_external_accounts():
    """Đọc danh sách tài khoản từ file JSON riêng biệt"""
    # Đảm bảo thư mục data tồn tại
    os.makedirs(os.path.dirname(ACCOUNTS_FILE), exist_ok=True)
    
    if not os.path.exists(ACCOUNTS_FILE):
        # Nếu chưa có file, khởi tạo mặc định 1 tài khoản Admin tối cao
        default_data = [
            {
                "username": "admin",
                "full_name": "Quản trị viên Hệ thống",
                "password": "123456",
                "role": "Quản trị viên",
                "is_active": True
            }
        ]
        save_external_accounts(default_data)
        return default_data
        
    try:
        with open(ACCOUNTS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []

def save_external_accounts(accounts_list):
    """Lưu danh sách tài khoản vào file JSON"""
    os.makedirs(os.path.dirname(ACCOUNTS_FILE), exist_ok=True)
    with open(ACCOUNTS_FILE, "w", encoding="utf-8") as f:
        json.dump(accounts_list, f, ensure_ascii=False, indent=4)

def verify_external_login(username, password):
    """Kiểm tra đăng nhập từ file JSON độc lập"""
    accounts = load_external_accounts()
    for acc in accounts:
        if acc["username"] == username and acc["password"] == password and acc.get("is_active", True):
            return acc
    return None

def sync_account_to_json(username, full_name, password, role, is_active):
    """Thêm mới hoặc Cập nhật thông tin tài khoản vào file JSON"""
    accounts = load_external_accounts()
    found = False
    
    for acc in accounts:
        if acc["username"] == username:
            acc["full_name"] = full_name
            if password: # Nếu có nhập mật khẩu mới thì mới cập nhật
                acc["password"] = password
            acc["role"] = role
            acc["is_active"] = is_active
            found = True
            break
            
    # Nếu là tài khoản mới tinh chưa có trong JSON
    if not found:
        accounts.append({
            "username": username,
            "full_name": full_name,
            "password": password,
            "role": role,
            "is_active": is_active
        })
        
    save_external_accounts(accounts)

def remove_account_from_json(username):
    """Xóa tài khoản khỏi file JSON"""
    accounts = load_external_accounts()
    # Lọc giữ lại các tài khoản khác với tên đăng nhập cần xóa
    accounts = [acc for acc in accounts if acc["username"] != username]
    save_external_accounts(accounts)