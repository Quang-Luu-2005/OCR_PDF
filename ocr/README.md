# OCR PDF Tool

Tool này nhận PDF scan hoặc digital, sửa lỗi OCR tiếng Anh, dịch sang tiếng Việt khoa học và xuất Markdown/Word có ảnh đúng vị trí.

## Cài đặt

Từ thư mục gốc repository:

```powershell
python -m pip install -r ocr/requirements.txt
Copy-Item ocr/.env.example ocr/.env
```

Mở `ocr/.env` và điền `GEMINI_API_KEY` lấy từ Google AI Studio. Dịch tiếng Việt được bật mặc định, dùng API tương thích OpenAI của Google tại `https://generativelanguage.googleapis.com/v1beta/openai/` với model `gemini-3.8-flash`; model không được tải về máy.

## Chạy một PDF

```powershell
cd ocr
python main.py --input input/caythuoc.pdf --output output --mode auto
```

Các mode:

- `auto`: tự nhận diện PDF scan hay digital.
- `scan`: ép chạy OCR.
- `digital`: ép xử lý PDF digital.

Khi dịch được bật, cả hai mode đều dùng Marker để tạo Markdown có cấu trúc và ánh xạ ảnh. Riêng mode `digital` vẫn giữ thêm file Word gốc do `pdf2docx` tạo.

Các đầu ra dịch:

- `<tên>_ocr_results.md`: tiếng Anh đã sửa lỗi OCR/chính tả.
- `<tên>_ocr_results_vi.md`: bản dịch tiếng Việt khoa học.
- `<tên>_ocr_results_vi.docx`: Word tiếng Việt với ảnh ở đúng khối Markdown.

Tắt dịch cho tài liệu vốn bằng tiếng Việt hoặc khi không muốn gọi API:

```powershell
python main.py --input input/tailieu.pdf --output output --no-translate-vi
```

Có thể đổi nhà cung cấp/model tương thích OpenAI mà không truyền API key qua CLI:

```powershell
python main.py --input input/paper.pdf --translation-model MODEL --translation-base-url https://example.com/v1
```

Đặt khóa của nhà cung cấp thay thế trong `TRANSLATION_API_KEY`. Checkpoint dịch nằm trong `ocr/artifacts/translation/`; chạy lại cùng nội dung sẽ tái sử dụng các chunk đã hoàn thành.

## Chạy hàng loạt

Đặt các PDF cần xử lý vào `ocr/input/`, sau đó chạy:

```powershell
cd ocr
python main.py --input input --output output --batch
```

Kết quả được lưu trong `ocr/output/`. Thư mục input/output và artifact trung gian là local, không commit lên Git.

## Notebook

Notebook thử nghiệm nằm tại [`notebooks/main.ipynb`](notebooks/main.ipynb). Cấu hình mặc định nằm trong [`config.py`](config.py).
