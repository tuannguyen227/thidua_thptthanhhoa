import os
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker, declarative_base
from contextlib import contextmanager

# 1. TỰ ĐỘNG LẤY LINK CƠ SỞ DỮ LIỆU ĐÁM MÂY (SUPABASE)
# Nếu đang code trên máy tính chưa có link, nó sẽ dùng tạm file SQLite
DB_URL = os.environ.get("DATABASE_URL", "sqlite:///data/thi_dua.db")

# Sửa lại tiền tố postgres:// thành postgresql:// để tương thích với các thư viện mới nhất
if DB_URL.startswith("postgres://"):
    DB_URL = DB_URL.replace("postgres://", "postgresql://", 1)

engine = create_engine(
    DB_URL, 
    echo=False,
    pool_pre_ping=True,      # Bắt hệ thống "gõ cửa" kiểm tra kết nối trước mỗi lần lấy dữ liệu
    pool_recycle=1800        # Tự động làm mới đường truyền mỗi 30 phút để tránh bị Cloud đá văng
)

SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

def get_session():
    return SessionLocal()

@contextmanager
def session_scope():
    """Cung cấp một scope giao dịch an toàn cho các thao tác CSDL."""
    session = get_session()
    try:
        yield session
        session.commit()
    except Exception:
        session.rollback()
        raise
    finally:
        session.close()

def seed_violation_categories():
    """Tự động thêm một số lỗi vi phạm mẫu nếu ngân hàng lỗi đang trống"""
    from database.models import ViolationCategory
    try:
        with session_scope() as session:
            if session.query(ViolationCategory).count() == 0:
                sample_errors = [
                    ViolationCategory(name="Đi học muộn", penalty_points=2.0),
                    ViolationCategory(name="Sai đồng phục", penalty_points=1.0),
                    ViolationCategory(name="Vắng không phép", penalty_points=5.0),
                    ViolationCategory(name="Mất trật tự", penalty_points=2.0)
                ]
                session.add_all(sample_errors)
    except Exception as e:
        print(f"Lỗi mồi dữ liệu ngân hàng lỗi: {e}")

def init_db():
    from database.models import Base, ViolationCategory, WeeklyViolation 
    Base.metadata.create_all(bind=engine)
    seed_violation_categories()