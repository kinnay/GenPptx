
FROM batonogov/pyinstaller-windows:latest AS builder

WORKDIR /temp

RUN pip install pillow python-pptx

COPY main.py .

RUN pyinstaller --onefile main.py

FROM scratch
COPY --from=builder /temp/dist/main.exe /GenPptx.exe
