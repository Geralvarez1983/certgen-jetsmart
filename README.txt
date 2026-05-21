╔══════════════════════════════════════════════════╗
║         CertGen – JetSMART  v1.0                ║
║    Generador de Certificados en Lote            ║
╚══════════════════════════════════════════════════╝

INSTALACIÓN (una sola vez)
──────────────────────────
1. Descomprimí esta carpeta donde quieras (ej: C:\CertGen)
2. Instalá Python 3.10+ desde https://python.org (marcá "Add to PATH")
3. Instalá LibreOffice desde https://www.libreoffice.org (necesario para PDF)
4. Abrí una terminal en la carpeta y ejecutá:
     pip install flask openpyxl

INICIAR LA APP
──────────────
Doble clic en  → INICIAR.bat  (Windows)
                 o ejecutá:  python app.py

Luego abrí el navegador en:  http://localhost:5000

CÓMO USARLA
───────────
1. MODELOS → Creá un modelo por tipo de certificado
   - Subís el .docx template
   - Subís la imagen de firma del instructor
   
2. GENERAR → Para emitir un lote:
   - Elegís el modelo
   - Subís el Excel con la nómina
   - Configurás qué columna es folio / nombre / DNI
   - Ingresás el folio inicial
   - ¡Generá! → Descargás el ZIP con todos los PDFs

PLACEHOLDERS EN EL WORD
────────────────────────
El Word debe contener estos textos exactos:
  • NOMBRE ALUMNO, DNI XXXXX  → se reemplaza por el piloto
  • DD/MM/YYYY                → fecha de cursada (opcional)
  • XX  (en el encabezado)    → número de folio

