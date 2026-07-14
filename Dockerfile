
FROM batonogov/pyinstaller-windows:latest AS builder

WORKDIR /temp
COPY main.py .

RUN pip install pillow python-pptx

RUN pyinstaller --onefile main.py

FROM scratch
COPY --from=builder /temp/dist/main.exe /GenPptx.exe
