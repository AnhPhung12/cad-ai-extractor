from motor.motor_asyncio import AsyncIOMotorClient
import os

# Chuỗi kết nối lấy từ trang Dashboard của MongoDB Atlas
# Định dạng: mongodb+srv://<username>:<password>@cluster0.xxxx.mongodb.net/?retryWrites=true&w=majority
MONGO_URI = os.getenv("MONGO_URI", "mongodb+srv://admin:your_password@cluster0.xxxxx.mongodb.net/?retryWrites=true&w=majority")
DB_NAME = "cad_bom_db"

client = None
mongo_db = None

async def connect_to_mongo():
    global client, mongo_db
    client = AsyncIOMotorClient(MONGO_URI)
    mongo_db = client[DB_NAME]
    print("[MongoDB Atlas] Kết nối thành công!")

async def close_mongo_connection():
    global client
    if client:
        client.close()
        print("[MongoDB Atlas] Đã ngắt kết nối.")

def get_mongo_db():
    return mongo_db