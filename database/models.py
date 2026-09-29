from sqlalchemy import Column, Integer, String, Boolean, ForeignKey, DateTime, Float, Enum, Text, Date, UniqueConstraint
from sqlalchemy.orm import relationship
from datetime import datetime, timedelta # [NÂNG CẤP]: Bổ sung timedelta để tính toán hạn chót 3 ngày
import enum
from .database import Base

# --- ENUMS ---
class UserRole(enum.Enum):
    ADMIN = "Quản trị viên"
    BI_THU = "Bí thư Đoàn trường"
    BCH = "BCH Đoàn trường"
    GVCN = "Giáo viên chủ nhiệm"
    SAO_DO = "Sao đỏ"

class ThiDuaGroup(enum.Enum):
    NHOM_1 = "Nhóm 1" 
    NHOM_2 = "Nhóm 2" 

# --- BẢNG: QUẢN LÝ KHÓA CHỐT SỔ ---
class PeriodLock(Base):
    __tablename__ = 'period_locks'
    id = Column(Integer, primary_key=True, autoincrement=True)
    period_name = Column(String, unique=True, nullable=False) 
    is_locked = Column(Boolean, default=False)

# --- MODELS ---
class User(Base):
    __tablename__ = "users"
    id = Column(Integer, primary_key=True, index=True)
    username = Column(String, unique=True, index=True)
    password_hash = Column(String) 
    full_name = Column(String)
    role = Column(Enum(UserRole), default=UserRole.BCH)
    is_active = Column(Boolean, default=True)

class SchoolYear(Base):
    __tablename__ = "school_years"
    id = Column(Integer, primary_key=True, index=True)
    name = Column(String, unique=True) 
    is_active = Column(Boolean, default=False)
    
    branches = relationship('Branch', back_populates='school_year', cascade='all, delete-orphan')

class Branch(Base):
    __tablename__ = "branches"
    id = Column(Integer, primary_key=True, index=True)
    name = Column(String) 
    group = Column(String)
    si_so = Column(Integer, default=0) 
    gvcn = Column(String)              
    phone_gvcn = Column(String)       # Số điện thoại GVCN
    class_monitor = Column(String)    # Tên Lớp trưởng
    phone_monitor = Column(String)    # Số điện thoại lớp trưởng
    school_year_id = Column(Integer, ForeignKey("school_years.id"))
    
    school_year = relationship("SchoolYear", back_populates="branches")
    red_stars = relationship('RedStar', back_populates='branch', cascade='all, delete-orphan')
    
    # Bổ sung cascade để xóa hết điểm tuần khi xóa năm học
    weekly_scores = relationship("WeeklyScore", back_populates="branch", cascade='all, delete-orphan')

class RedStar(Base):
    __tablename__ = 'red_stars'
    id = Column(Integer, primary_key=True, autoincrement=True)
    full_name = Column(String(100), nullable=False)
    gender = Column(String(10), default="Nam")    
    phone = Column(String(20))
    branch_id = Column(Integer, ForeignKey('branches.id'), nullable=False)
    notes = Column(String(255), default="")       
    is_active = Column(Boolean, default=True)
    
    branch = relationship("Branch", back_populates="red_stars")
    
    # Bổ sung cascade cho phân công trực
    assignments = relationship("Assignment", back_populates="red_star", cascade='all, delete-orphan')
    
    # Bổ sung dòng này để quy định rõ việc xóa sổ đánh giá khi xóa năm học
    evaluations = relationship("StarEvaluation", back_populates="red_star", cascade="all, delete-orphan")

class DutyArea(Base):
    __tablename__ = "duty_areas"
    id = Column(Integer, primary_key=True, index=True)
    name = Column(String) 
    required_stars = Column(Integer, default=2) 

class Assignment(Base):
    __tablename__ = "assignments"
    id = Column(Integer, primary_key=True, index=True)
    red_star_id = Column(Integer, ForeignKey("red_stars.id"))
    area_id = Column(Integer, ForeignKey("duty_areas.id"))
    week_number = Column(Integer)
    date = Column(DateTime)
    shift = Column(String) 
    
    red_star = relationship("RedStar", back_populates="assignments")
    duty_area = relationship("DutyArea")

class DutyRecord(Base):
    __tablename__ = "duty_records"
    id = Column(Integer, primary_key=True, index=True)
    branch_id = Column(Integer, ForeignKey("branches.id"))
    week_number = Column(Integer) 
    tong_diem_sao_do = Column(Float, default=100.0) 
    diem_tot = Column(Integer, default=0) 
    diem_kem = Column(Integer, default=0) 
    diem_cong = Column(Float, default=0.0) 
    diem_tru = Column(Float, default=0.0) 
    chi_tiet_vi_pham = Column(String) 
    created_at = Column(DateTime, default=datetime.utcnow)

class WeeklyScore(Base):
    __tablename__ = "weekly_scores"
    id = Column(Integer, primary_key=True, index=True)
    branch_id = Column(Integer, ForeignKey('branches.id'))
    week = Column(String)  
    score_truc = Column(Integer, default=0)
    score_tot = Column(Integer, default=0)
    score_kem = Column(Integer, default=0)
    score_cong = Column(Integer, default=0)
    score_tru = Column(Integer, default=0)
    note = Column(String)
    total_score = Column(Float, default=100.0)
    rank = Column(Integer)
    is_locked = Column(Boolean, default=False)
    start_date = Column(String, nullable=True)
    end_date = Column(String, nullable=True)
    
    # --- CƠ CHẾ KHIẾU NẠI DÀNH CHO GVCN ---
    is_appealed = Column(Boolean, default=False)       
    appeal_reason = Column(String(255), nullable=True) 
    appeal_response = Column(String(255), nullable=True) 
    
    # --- CỘT BỔ SUNG CHO TÍNH NĂNG NHẬP ĐIỂM CHUYÊN SÂU ---
    count_8 = Column(Integer, default=0)
    count_9 = Column(Integer, default=0)
    count_10 = Column(Integer, default=0)
    week_rating = Column(String, default="Bình thường")
    note_cong = Column(String, nullable=True)
    note_tru = Column(String, nullable=True)
    
    evidence_image = Column(String(255), nullable=True)
    # [NÂNG CẤP LÕI]: Cột lưu mốc thời gian chấm điểm để giới hạn 3 ngày phúc khảo
    created_at = Column(DateTime, default=datetime.utcnow)
    # -----------------------------------------------------

    branch = relationship("Branch", back_populates="weekly_scores")
    violations = relationship("WeeklyViolation", back_populates="weekly_score", cascade="all, delete-orphan")

    # [NÂNG CẤP LÕI]: Hàm ảo kiểm tra xem bảng điểm này đã quá hạn 3 ngày chưa
    @property
    def is_appeal_expired(self):
        if self.created_at:
            # Cộng thêm 3 ngày vào thời điểm tạo bảng điểm
            deadline = self.created_at + timedelta(days=3)
            # So sánh thời gian hiện tại với deadline
            return datetime.utcnow() > deadline
        return False # Trả về False nếu là dữ liệu cũ chưa có created_at để tránh lỗi

class MonthlyRecord(Base):
    __tablename__ = 'monthly_records'
    id = Column(Integer, primary_key=True, autoincrement=True)
    school_year_id = Column(Integer, ForeignKey('school_years.id'), nullable=False)
    month_name = Column(String, nullable=False)   
    branch_id = Column(Integer, ForeignKey('branches.id'), nullable=False)
    total_score = Column(Float, default=0.0)      
    rank = Column(Integer, nullable=False)        
    weeks_used = Column(String, nullable=True)    
    start_date = Column(String, nullable=True)
    end_date = Column(String, nullable=True)   
    
class StarEvaluation(Base):
    __tablename__ = 'star_evaluations'
    id = Column(Integer, primary_key=True)
    
    # Thông tin người đánh giá (Ai đang chấm điểm)
    evaluator_username = Column(String(50), nullable=False)
    evaluator_role = Column(String(50), nullable=False)
    
    # Thông tin người bị đánh giá (Sao đỏ nào bị chấm)
    evaluatee_id = Column(Integer, ForeignKey('red_stars.id'), nullable=False)
    week_name = Column(String(20), nullable=False)
    
    # 4 Tiêu chí trắc nghiệm (Lưu giá trị: 1, 3 hoặc 5)
    score_gio_giac = Column(Integer, default=5)
    score_tac_phong = Column(Integer, default=5)
    score_thai_do = Column(Integer, default=5)
    score_cong_tam = Column(Integer, default=5)
    
    comment = Column(Text, nullable=True) # Ý kiến góp ý ẩn danh
    
    # ĐÃ SỬA: Xóa bỏ backref="evaluations" để tránh xung đột
    red_star = relationship("RedStar")

class ActionLog(Base):
    __tablename__ = "action_logs"
    id = Column(Integer, primary_key=True, autoincrement=True)
    username = Column(String, nullable=False)      
    full_name = Column(String, nullable=True)     
    action_type = Column(String, nullable=False)  
    details = Column(String, nullable=True)       
    timestamp = Column(DateTime, default=datetime.utcnow)

# ==========================================
# CÁC BẢNG CHO NGÂN HÀNG TIÊU CHÍ VÀ LỖI
# ==========================================
class ViolationCategory(Base):
    __tablename__ = 'violation_categories'
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    name = Column(String(255), nullable=False)
    penalty_points = Column(Float, default=0.0)
    point_type = Column(String(50), default="Điểm trừ")
    
    weekly_violations = relationship("WeeklyViolation", back_populates="category", cascade="all, delete-orphan")
    
    school_year_id = Column(Integer, ForeignKey('school_years.id'))
    school_year = relationship("SchoolYear")

class WeeklyViolation(Base):
    __tablename__ = 'weekly_violations'
    id = Column(Integer, primary_key=True, autoincrement=True)
    weekly_score_id = Column(Integer, ForeignKey('weekly_scores.id', ondelete="CASCADE"))
    violation_id = Column(Integer, ForeignKey('violation_categories.id'))
    quantity = Column(Integer, default=1) 
    
    # --- THEO DÕI SỔ ĐEN CÁ NHÂN ---
    student_name = Column(String(255), nullable=True)  
    # ----------------------------------------
    
    weekly_score = relationship("WeeklyScore", back_populates="violations")
    category = relationship("ViolationCategory", back_populates="weekly_violations")

# --- BẢNG CẤU HÌNH BAREM ĐIỂM THEO NĂM HỌC ---
class ScoreSettings(Base):
    __tablename__ = 'score_settings'
    id = Column(Integer, primary_key=True, autoincrement=True)
    school_year_id = Column(Integer, ForeignKey('school_years.id', ondelete="CASCADE"), unique=True, nullable=False)
    
    diem_8 = Column(Float, default=1.0)
    diem_9 = Column(Float, default=3.0)
    diem_10 = Column(Float, default=5.0)
    diem_tuan_tot = Column(Float, default=30.0)
    diem_tuan_kha = Column(Float, default=20.0)
    max_diem_tot = Column(Integer, default=14.0)
    max_diem_mon = Column(Float, default=4.0)

class RawScore(Base):
    __tablename__ = 'raw_scores'
    
    id = Column(Integer, primary_key=True)
    week = Column(String)           # Lưu tên Tuần (VD: "Tuần 1")
    branch_name = Column(String)    # Lưu tên Chi đoàn (VD: "10A1")
    subject = Column(String)        # Tên môn học (VD: "Toán", "Lý")
    c10 = Column(Integer, default=0)
    c9 = Column(Integer, default=0)
    c8 = Column(Integer, default=0)

class ScoreAuditLog(Base):
    __tablename__ = 'score_audit_logs'
    id = Column(Integer, primary_key=True, autoincrement=True)
    timestamp = Column(DateTime, default=datetime.utcnow)
    username = Column(String(50))
    full_name = Column(String(100))
    branch_name = Column(String(50))
    week = Column(String(20))
    old_score = Column(Float)
    new_score = Column(Float)
    details = Column(String(255))

# ==========================================
# [NÂNG CẤP MỚI]: BẢNG ĐIỂM DANH GVCN
# ==========================================
class GVCNAttendance(Base):
    __tablename__ = 'gvcn_attendances'
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    branch_id = Column(Integer, ForeignKey('branches.id', ondelete='CASCADE'), nullable=False)
    week_name = Column(String(50), nullable=False)
    date = Column(Date, nullable=False)
    created_at = Column(DateTime, default=datetime.utcnow)

    # Liên kết ngược (relationship) để dễ dàng truy vấn từ Lớp học
    branch = relationship('Branch', backref='attendances')
    
    # Ràng buộc cấp cơ sở dữ liệu: Mỗi lớp 1 ngày chỉ được ghi nhận 1 lần duy nhất
    __table_args__ = (
        UniqueConstraint('branch_id', 'date', name='unique_attendance_per_day'),
    )