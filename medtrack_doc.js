const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak, LevelFormat,
  TableOfContents, ExternalHyperlink
} = require('docx');
const fs = require('fs');

// ─── helpers ───────────────────────────────────────────────────────────────
const BLUE      = '185FA5';
const BLUE_DARK = '0C447C';
const BLUE_LIGHT= 'D6E9F8';
const GRAY_BG   = 'F2F4F7';
const GRAY_DARK = '374151';
const GREEN     = '0F6E56';
const GREEN_BG  = 'D1FAE5';
const RED       = 'A32D2D';
const RED_BG    = 'FEE2E2';
const AMBER_BG  = 'FEF3C7';
const AMBER     = '854F0B';
const WHITE     = 'FFFFFF';

const border = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

const CONTENT_W = 9360; // US Letter, 1" margins

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 400, after: 160 },
    children: [new TextRun({ text, font: 'Arial', size: 36, bold: true, color: BLUE_DARK })]
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 320, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BLUE, space: 4 } },
    children: [new TextRun({ text, font: 'Arial', size: 28, bold: true, color: BLUE })]
  });
}
function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 240, after: 80 },
    children: [new TextRun({ text, font: 'Arial', size: 24, bold: true, color: GRAY_DARK })]
  });
}
function h4(text) {
  return new Paragraph({
    spacing: { before: 200, after: 60 },
    children: [new TextRun({ text, font: 'Arial', size: 22, bold: true, color: '4B5563' })]
  });
}
function para(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 80, after: 80 },
    children: [new TextRun({ text, font: 'Arial', size: 22, color: GRAY_DARK, ...opts })]
  });
}
function paraRuns(runs, spacing = {}) {
  return new Paragraph({
    spacing: { before: 80, after: 80, ...spacing },
    children: runs.map(r => new TextRun({ font: 'Arial', size: 22, color: GRAY_DARK, ...r }))
  });
}
function bullet(text, level = 0) {
  return new Paragraph({
    numbering: { reference: 'bullets', level },
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text, font: 'Arial', size: 22, color: GRAY_DARK })]
  });
}
function numbered(text, level = 0) {
  return new Paragraph({
    numbering: { reference: 'numbers', level },
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text, font: 'Arial', size: 22, color: GRAY_DARK })]
  });
}
function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}
function spacer(lines = 1) {
  return new Paragraph({ spacing: { before: 0, after: lines * 120 }, children: [new TextRun('')] });
}
function infoBox(text, bgColor, textColor = GRAY_DARK) {
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W],
    rows: [new TableRow({ children: [new TableCell({
      borders: noBorders,
      shading: { fill: bgColor, type: ShadingType.CLEAR },
      margins: { top: 120, bottom: 120, left: 160, right: 160 },
      width: { size: CONTENT_W, type: WidthType.DXA },
      children: [new Paragraph({ children: [new TextRun({ text, font: 'Arial', size: 21, color: textColor })] })]
    })]})],
  });
}
function codeBlock(lines) {
  const rows = Array.isArray(lines) ? lines : [lines];
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W],
    rows: [new TableRow({ children: [new TableCell({
      borders: { top: border, bottom: border, left: { style: BorderStyle.SINGLE, size: 12, color: BLUE }, right: border },
      shading: { fill: '1E2937', type: ShadingType.CLEAR },
      margins: { top: 120, bottom: 120, left: 200, right: 160 },
      width: { size: CONTENT_W, type: WidthType.DXA },
      children: rows.map(line => new Paragraph({
        spacing: { before: 20, after: 20 },
        children: [new TextRun({ text: line, font: 'Courier New', size: 18, color: 'E2E8F0' })]
      }))
    })]})],
  });
}
function headerRow(cells, widths) {
  return new TableRow({
    tableHeader: true,
    children: cells.map((c, i) => new TableCell({
      borders,
      shading: { fill: BLUE, type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      width: { size: widths[i], type: WidthType.DXA },
      verticalAlign: VerticalAlign.CENTER,
      children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: c, font: 'Arial', size: 20, bold: true, color: WHITE })] })]
    }))
  });
}
function dataRow(cells, widths, shade = false) {
  return new TableRow({
    children: cells.map((c, i) => new TableCell({
      borders,
      shading: { fill: shade ? GRAY_BG : WHITE, type: ShadingType.CLEAR },
      margins: { top: 70, bottom: 70, left: 120, right: 120 },
      width: { size: widths[i], type: WidthType.DXA },
      children: [new Paragraph({ children: [new TextRun({ text: c, font: 'Arial', size: 20, color: GRAY_DARK })] })]
    }))
  });
}
function table(headerCells, rows, widths) {
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: widths,
    rows: [
      headerRow(headerCells, widths),
      ...rows.map((r, i) => dataRow(r, widths, i % 2 === 1))
    ]
  });
}

// ─── DOCUMENT ──────────────────────────────────────────────────────────────
const doc = new Document({
  numbering: {
    config: [
      { reference: 'bullets', levels: [
        { level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
        { level: 1, format: LevelFormat.BULLET, text: '◦', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1080, hanging: 360 } } } },
      ]},
      { reference: 'numbers', levels: [
        { level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
        { level: 1, format: LevelFormat.LOWER_LETTER, text: '%2)', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1080, hanging: 360 } } } },
      ]},
    ]
  },
  styles: {
    default: { document: { run: { font: 'Arial', size: 22 } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 36, bold: true, font: 'Arial', color: BLUE_DARK },
        paragraph: { spacing: { before: 400, after: 160 }, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 28, bold: true, font: 'Arial', color: BLUE },
        paragraph: { spacing: { before: 320, after: 120 }, outlineLevel: 1 } },
      { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 24, bold: true, font: 'Arial', color: GRAY_DARK },
        paragraph: { spacing: { before: 240, after: 80 }, outlineLevel: 2 } },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    headers: {
      default: new Header({
        children: [
          new Table({
            width: { size: CONTENT_W, type: WidthType.DXA },
            columnWidths: [5000, 4360],
            rows: [new TableRow({ children: [
              new TableCell({
                borders: { ...noBorders, bottom: { style: BorderStyle.SINGLE, size: 4, color: BLUE } },
                width: { size: 5000, type: WidthType.DXA },
                children: [new Paragraph({ children: [new TextRun({ text: 'MedTrack — Teknik Geliştirme Dokümantasyonu', font: 'Arial', size: 18, bold: true, color: BLUE })] })]
              }),
              new TableCell({
                borders: { ...noBorders, bottom: { style: BorderStyle.SINGLE, size: 4, color: BLUE } },
                width: { size: 4360, type: WidthType.DXA },
                children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: 'v1.0  |  Gizli', font: 'Arial', size: 18, color: '9CA3AF' })] })]
              })
            ]})],
          })
        ]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'E5E7EB', space: 4 } },
          children: [
            new TextRun({ text: 'MedTrack Akıllı Hastane Sistemi  |  Sayfa ', font: 'Arial', size: 18, color: '9CA3AF' }),
            new TextRun({ children: [new PageNumber()], font: 'Arial', size: 18, color: '9CA3AF' }),
          ]
        })]
      })
    },
    children: [

      // ════════════════════════════════════════
      // KAPAK SAYFASI
      // ════════════════════════════════════════
      new Table({
        width: { size: CONTENT_W, type: WidthType.DXA },
        columnWidths: [CONTENT_W],
        rows: [new TableRow({ children: [new TableCell({
          borders: noBorders,
          shading: { fill: BLUE, type: ShadingType.CLEAR },
          margins: { top: 600, bottom: 600, left: 400, right: 400 },
          width: { size: CONTENT_W, type: WidthType.DXA },
          children: [
            new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 200, after: 60 }, children: [new TextRun({ text: '🏥  MedTrack', font: 'Arial', size: 64, bold: true, color: WHITE })] }),
            new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 80 }, children: [new TextRun({ text: 'Akıllı Hastane Hasta Takip &', font: 'Arial', size: 36, color: 'B3D4F5' })] }),
            new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 200 }, children: [new TextRun({ text: 'İlaç Doğrulama Sistemi', font: 'Arial', size: 36, color: 'B3D4F5' })] }),
            new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 60 }, children: [new TextRun({ text: 'TEKNİK GELİŞTİRME DOKÜMANTASYONU', font: 'Arial', size: 22, bold: true, color: '93C5FD' })] }),
            new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: 'Versiyon 1.0  —  NFC + RFID Tabanlı Hastane Yönetim Platformu', font: 'Arial', size: 20, color: '93C5FD' })] }),
          ]
        })]})],
      }),
      spacer(2),
      new Table({
        width: { size: CONTENT_W, type: WidthType.DXA },
        columnWidths: [3120, 3120, 3120],
        rows: [new TableRow({ children: [
          ['Hazırlayan', 'Proje Adı', 'Tarih'].map((label, i) => {
            const vals = ['Geliştirici Ekibi', 'MedTrack v1.0', '2025'];
            return new TableCell({
              borders: { ...noBorders, top: { style: BorderStyle.SINGLE, size: 6, color: BLUE } },
              width: { size: 3120, type: WidthType.DXA },
              margins: { top: 120, bottom: 0, left: 0, right: 0 },
              children: [
                new Paragraph({ children: [new TextRun({ text: label, font: 'Arial', size: 18, color: '9CA3AF' })] }),
                new Paragraph({ children: [new TextRun({ text: vals[i], font: 'Arial', size: 22, bold: true, color: BLUE_DARK })] }),
              ]
            });
          })
        ]})],
      }),
      pageBreak(),

      // ════════════════════════════════════════
      // İÇİNDEKİLER
      // ════════════════════════════════════════
      h1('İçindekiler'),
      new TableOfContents('İçindekiler', {
        hyperlink: true, headingStyleRange: '1-3',
        stylesWithLevels: [
          { styleName: 'Heading 1', level: 1 },
          { styleName: 'Heading 2', level: 2 },
          { styleName: 'Heading 3', level: 3 },
        ]
      }),
      pageBreak(),

      // ════════════════════════════════════════
      // 1. PROJE GENEL BAKIŞ
      // ════════════════════════════════════════
      h1('1. Proje Genel Bakış'),
      h2('1.1  Proje Özeti'),
      para('MedTrack, NFC (Near Field Communication) ve RFID (Radio Frequency Identification) teknolojilerini birleştiren, hastane ortamlarında hasta kimlik doğrulama, ilaç güvenliği ve ekipman takibini tek platformda sunan web tabanlı bir yönetim sistemidir.'),
      spacer(),
      para('Sistem iki temel sorunu çözmektedir:'),
      bullet('Yanlış ilaç uygulaması: Dünya genelinde tıbbi hataların %26\'sı yanlış ilaç veya yanlış doza bağlıdır. MedTrack, 5 aşamalı RFID doğrulamasıyla bu riski sıfıra yakın indirir.'),
      bullet('Hasta takibi: NFC bilezikler sayesinde hasta kimliği, konum ve tıbbi geçmiş anlık sorgulanabilir; yetkisiz bölge girişleri otomatik alarm tetikler.'),
      spacer(),
      infoBox('Kapsam: Bu doküman, MedTrack web uygulamasının tam teknik geliştirme spesifikasyonunu içerir. Frontend, backend, veritabanı şeması, API sözleşmeleri, güvenlik gereksinimleri ve test senaryolarını kapsar. Dokümanı okuyan geliştirici, herhangi bir ek bilgeye ihtiyaç duymaksızın sistemi sıfırdan inşa edebilmelidir.', BLUE_LIGHT, BLUE_DARK),
      spacer(),

      h2('1.2  Teknoloji Çifti: NFC + RFID'),
      table(
        ['Teknoloji', 'Kullanım Amacı', 'Kapsam', 'Frekans', 'Neden Bu Proje?'],
        [
          ['NFC', 'Hasta kimlik doğrulama', '≤ 10 cm', '13.56 MHz', 'Hasta bileziklerinden anlık kimlik okuma; Kart Emulyasyonu modu'],
          ['RFID (HF)', 'İlaç kutu etiketleri', '≤ 10 cm', '13.56 MHz', 'Her ilaç kutusuna yapıştırılan RFID etiketiyle doğrulama'],
          ['RFID (UHF)', 'Ekipman/hasta konum takibi', '1–10 m', '860–960 MHz', 'Koğuş antenlerinden pasif konum tespiti'],
        ],
        [2200, 2200, 1400, 1400, 2160]
      ),
      spacer(),

      h2('1.3  Temel Kullanıcı Rolleri'),
      table(
        ['Rol', 'Kimler?', 'Yetkiler', 'NFC/RFID Etkileşimi'],
        [
          ['Süper Admin', 'BT Yöneticisi', 'Tüm sistem + kullanıcı yönetimi + raporlar', 'Okuyucu yapılandırması'],
          ['Doktor', 'Sorumlu Hekim', 'Hasta dosyası okuma/yazma, ilaç atama', 'NFC kimlik doğrulama'],
          ['Hemşire', 'Koğuş Hemşiresi', 'İlaç doğrulama, vital kayıt, hasta giriş/çıkış', 'NFC + RFID tarama'],
          ['Eczacı', 'Hastane Eczacısı', 'İlaç stok, RFID etiket atama, uyarı yönetimi', 'RFID etiket yazma'],
          ['Güvenlik', 'Güvenlik Görevlisi', 'Konum haritası görüntüleme, alarm onaylama', 'Sadece okuma'],
        ],
        [1600, 1800, 3360, 2600]
      ),
      spacer(),

      h2('1.4  Yüksek Seviye Mimari'),
      para('MedTrack üç katmanlı bir mimari üzerinde çalışır:'),
      spacer(),
      bullet('Sunum Katmanı (Frontend): React.js tabanlı SPA. Tüm modüller (hasta, ilaç, konum, alarmlar) tek sayfalı uygulama içinde yönetilir. WebSocket bağlantısı ile gerçek zamanlı güncelleme alır.'),
      bullet('İş Mantığı Katmanı (Backend): Node.js + Express REST API. JWT kimlik doğrulama, rol bazlı yetkilendirme, ilaç doğrulama motoru ve WebSocket sunucusu bu katmanda çalışır.'),
      bullet('Veri Katmanı: PostgreSQL ana veritabanı + Redis önbelleği. RFID/NFC okuma logları yüksek yazma hızına sahip olduğundan PostgreSQL\'e async yazılır; sık sorgulanan veriler Redis\'te cache\'lenir.'),
      bullet('Donanım Köprüsü: Fiziksel RFID/NFC okuyucular (ACR122U veya benzer) USB/WebSerial API veya yerel bir agent servisi aracılığıyla backend\'e bağlanır.'),
      pageBreak(),

      // ════════════════════════════════════════
      // 2. FONKSİYONEL GEREKSİNİMLER
      // ════════════════════════════════════════
      h1('2. Fonksiyonel Gereksinimler'),

      h2('2.1  Modül 1 — Hasta Yönetimi (NFC Kimlik)'),
      h3('2.1.1  Hasta Kaydı'),
      para('Her hasta sisteme kaydedildiğinde aşağıdaki veriler girilir ve bir NFC bilezik UID\'i atanır:'),
      table(
        ['Alan', 'Tip', 'Zorunlu', 'Açıklama'],
        [
          ['ad_soyad', 'VARCHAR(120)', 'Evet', 'Tam ad — TC Kimlik ile eşleşmeli'],
          ['tc_kimlik', 'CHAR(11)', 'Evet', 'Şifreli saklanır, görüntülenmez'],
          ['dogum_tarihi', 'DATE', 'Evet', 'Yaş hesaplamasında kullanılır'],
          ['cinsiyet', 'ENUM(E,K,Belirtilmemiş)', 'Evet', '—'],
          ['kan_grubu', 'VARCHAR(5)', 'Hayır', 'A+, B-, O+ vb.'],
          ['alerjiler', 'TEXT[ ]', 'Hayır', 'Çoklu — ilaç doğrulamada kritik'],
          ['kronik_hastaliklar', 'TEXT[ ]', 'Hayır', 'Diyabet, hipertansiyon vb.'],
          ['sorumlu_doktor_id', 'UUID', 'Evet', 'FK → kullanicilar tablosu'],
          ['kogus_oda', 'VARCHAR(20)', 'Evet', 'A-201, B-104 formatı'],
          ['nfc_uid', 'VARCHAR(32)', 'Evet', 'Bilezik UID — benzersiz olmalı'],
          ['yatis_tarihi', 'TIMESTAMP', 'Evet', 'Otomatik set edilir'],
          ['durum', 'ENUM', 'Evet', 'Stabil / İzleme / Kritik / Taburcu'],
        ],
        [2000, 2000, 1400, 3960]
      ),
      spacer(),

      h3('2.1.2  NFC Kimlik Doğrulama Akışı'),
      para('Koğuş giriş noktalarına yerleştirilen NFC okuyucular aşağıdaki sırayla çalışır:'),
      numbered('Hemşire, NFC okuyucuyu aktive eder (web arayüzünden "Okuyucuyu Başlat" tuşu).'),
      numbered('Hasta bilezikten NFC UID okunur (≤ 0.5 saniye).'),
      numbered('Backend UID\'i veritabanında arar; eşleşme bulunursa hasta profili döndürülür.'),
      numbered('Hasta durumu "Kritik" ise kırmızı uyarı banner\'ı görüntülenir.'),
      numbered('Giriş kaydı: kimlik, zaman damgası, okuyucu ID, oda kodu audit_log tablosuna yazılır.'),
      numbered('NFC UID sistemde kayıtlı değilse "Bilinmeyen Bilezik" alarmı tetiklenir.'),
      spacer(),
      infoBox('Güvenlik Notu: NFC UID\'ler tek başına kimlik doğrulama için yeterli değildir. Sistem ek olarak okuyucunun fiziksel konumunu (hangi koğuşta olduğunu) doğrular. A Koğuşu okuyucusundan C Koğuşu hastasının UID\'i okunursa "Yanlış Koğuş" uyarısı verilir.', AMBER_BG, AMBER),
      spacer(),

      h2('2.2  Modül 2 — İlaç Doğrulama (RFID 5-Adım Motoru)'),
      para('Bu modül sistemin kalbidir. Bir hemşire ilacı hastaya vermeden önce RFID okuyucudan geçirmek zorundadır. Sistem beş bağımsız kontrolü sırayla çalıştırır:'),
      spacer(),
      table(
        ['Adım', 'Kontrol Adı', 'Mantık', 'Başarısız Sonuç'],
        [
          ['1', 'Hasta-İlaç Eşleşmesi', 'RFID UID\'i taranan ilacın, ilgili hastanın ilaç listesinde bulunup bulunmadığı kontrol edilir.', 'BLOKAJ — İlaç bu hastaya atanmamış'],
          ['2', 'Doz Uygunluğu', 'İlacın günlük doz sayısı kontrol edilir; bu gün kaç kez verildiğine bakılır.', 'BLOKAJ — Maksimum günlük doz aşıldı'],
          ['3', 'Alerji Çakışması', 'İlaçtaki etken maddeler hastanın alerji listesiyle kesiştirme (intersection) sorgusuyla karşılaştırılır.', 'BLOKAJ — Alerji çakışması tespit edildi'],
          ['4', 'Son Kullanma Tarihi', 'RFID etiketinde kayıtlı SKT, bugünün tarihiyle karşılaştırılır.', 'BLOKAJ — İlaç süresi dolmuş'],
          ['5', 'Çift Doz Engeli', 'Aynı ilaç bu hastaya son 4 saat içinde verilmiş mi kontrol edilir.', 'BLOKAJ — Çift doz riski'],
        ],
        [700, 2000, 3760, 2900]
      ),
      spacer(),
      para('Tüm kontroller geçildiğinde sistem "ONAYLANDI" yanıtı döndürür ve ilaç_log tablosuna kayıt atar. Herhangi bir kontrol başarısız olursa ilaç uygulaması kilitlenir; hemşire yalnızca sorumlu doktor onayıyla geçersiz kılabilir (override_log kaydı oluşur).'),
      spacer(),

      h3('2.2.1  RFID İlaç Etiketi Yapısı'),
      para('Her ilaç kutusuna yapıştırılan RFID etiketi (NTAG213, 144 byte) aşağıdaki yapıda veri taşır:'),
      codeBlock([
        '{',
        '  "uid": "RF:M1:4A:2B",          // Etiket benzersiz ID',
        '  "ilac_id": "MED-00147",         // Sistemdeki ilaç kaydı ID',
        '  "adi": "Metoprolol 50mg",',
        '  "etken_madde": ["Metoprolol tartrat"],',
        '  "lot_no": "LOT-2024-0892",',
        '  "skt": "2026-08",               // YYYY-MM formatı',
        '  "uretici": "Abdi İbrahim",',
        '  "atc_kodu": "C07AB02",          // WHO ATC kodu',
        '  "barkod": "8699525012345"',
        '}',
      ]),
      spacer(),

      h2('2.3  Modül 3 — Gerçek Zamanlı Konum Takibi'),
      para('Hastane koridorlarına ve koğuşlara yerleştirilen UHF RFID anten okuyucuları (sabit, 4 anten/okuyucu), hasta bileziklerindeki pasif UHF etiketleri 1–3 saniyede bir okur ve konum verisini günceller.'),
      spacer(),
      h4('Konum Belirleme Algoritması'),
      numbered('Her anten, etiketi algıladığında RSSI (sinyal gücü) değerini kaydeder.'),
      numbered('Backend, aynı etikete ait en yüksek RSSI\'yi veren anteni "aktif anten" seçer.'),
      numbered('Aktif antenin fiziksel konum bilgisi (bina, kat, oda kodu) hasta kaydına yazılır.'),
      numbered('Konum değişikliği tespit edildiğinde WebSocket üzerinden tüm aktif istemcilere broadcast yapılır.'),
      spacer(),
      h4('İzin Dışı Bölge Alarmı'),
      para('Admin panelinden her hasta için izin verilen bölgeler tanımlanabilir. Sistem, hastanın bu bölge dışına çıkmasını tespit ettiğinde:'),
      bullet('Anlık alarm kaydı oluşturur (seviye: YÜKSEK).'),
      bullet('Güvenlik görevlisinin ekranında popup bildirim gösterir.'),
      bullet('E-posta / SMS entegrasyonu yapılandırılmışsa bildirim gönderir.'),
      spacer(),

      h2('2.4  Modül 4 — Alarm Yönetimi'),
      table(
        ['Alarm Türü', 'Seviye', 'Tetikleyici', 'Otomatik Eylem'],
        [
          ['Alerji Çakışması', 'KRİTİK', 'İlaç doğrulama adım 3 başarısız', 'İlaç kilitli + Doktor bildirim'],
          ['Gecikmiş İlaç', 'YÜKSEK', 'Zamanı geçmiş doz (> 30 dk)', 'Hemşire + Doktor bildirim'],
          ['İzin Dışı Bölge', 'YÜKSEK', 'Hasta konum anteni değişimi', 'Güvenlik popup + SMS'],
          ['Bilinmeyen NFC', 'ORTA', 'Kayıtsız UID taraması', 'Güvenlik log kaydı'],
          ['SKT Geçmiş İlaç', 'ORTA', 'İlaç doğrulama adım 4 başarısız', 'Eczane bildirim'],
          ['Stok Uyarısı', 'DÜŞÜK', 'İlaç miktarı eşik altına düşme', 'Eczane e-posta'],
          ['Çift Doz', 'KRİTİK', 'İlaç doğrulama adım 5 başarısız', 'İlaç kilitli + doktor onayı gerekli'],
        ],
        [2000, 1300, 2800, 3260]
      ),
      spacer(),

      h2('2.5  Modül 5 — Raporlama & Audit Log'),
      para('Her işlem (NFC okuma, ilaç doğrulama, konum değişimi, alarm) değiştirilemez audit_log tablosuna kayıt edilir. Raporlar üç formatta dışa aktarılabilir: PDF, Excel (XLSX), CSV.'),
      bullet('Günlük İlaç Raporu: Hangi hastaya, hangi ilaç, hangi hemşire tarafından, ne zaman verildi.'),
      bullet('Erişim Raporu: Belirli bir zaman aralığında hangi NFC UID, hangi kapıdan geçti.'),
      bullet('Alarm Raporu: Alarm türü, sıklık analizi, çözüm süresi.'),
      bullet('Ekipman Raporu: RFID etiketli ekipmanların konum geçmişi.'),
      pageBreak(),

      // ════════════════════════════════════════
      // 3. TEKNİK MİMARİ
      // ════════════════════════════════════════
      h1('3. Teknik Mimari'),

      h2('3.1  Teknoloji Yığını'),
      table(
        ['Katman', 'Teknoloji', 'Versiyon', 'Sebep'],
        [
          ['Frontend', 'React.js', '18.x', 'Bileşen bazlı mimari, geniş ekosistem'],
          ['Frontend UI', 'Tailwind CSS', '3.x', 'Utility-first; hızlı prototipleme'],
          ['Frontend State', 'Zustand', '4.x', 'Redux\'a göre çok daha basit; async yok'],
          ['Gerçek Zamanlı', 'Socket.io (istemci)', '4.x', 'WebSocket + fallback'],
          ['Backend', 'Node.js + Express', '20.x LTS', 'Yüksek eşzamanlı bağlantı desteği'],
          ['Gerçek Zamanlı', 'Socket.io (sunucu)', '4.x', 'NFC/RFID olaylarını push eder'],
          ['Veritabanı', 'PostgreSQL', '16.x', 'ACID uyumlu; JSON sütunu desteği'],
          ['Önbellek', 'Redis', '7.x', 'Session, rate-limit, hasta önbelleği'],
          ['ORM', 'Prisma', '5.x', 'Tip güvenli sorgular; migration yönetimi'],
          ['Kimlik Doğrulama', 'JWT + bcrypt', '—', 'Stateless; rol bazlı yetkilendirme'],
          ['Donanım Köprüsü', 'node-nfc-nci / custom agent', '—', 'USB RFID okuyucu sürücüsü'],
          ['Container', 'Docker + Compose', '—', 'Ortam bağımsızlığı'],
          ['Proxy', 'Nginx', '1.25', 'SSL sonlandırma, load balancing'],
        ],
        [2000, 2200, 1400, 3760]
      ),
      spacer(),

      h2('3.2  Veritabanı Şeması'),
      h3('Ana Tablolar'),
      codeBlock([
        '-- Kullanıcılar (doktor, hemşire, eczacı, admin)',
        'CREATE TABLE kullanicilar (',
        '  id          UUID PRIMARY KEY DEFAULT gen_random_uuid(),',
        '  ad_soyad    VARCHAR(120) NOT NULL,',
        '  email       VARCHAR(180) UNIQUE NOT NULL,',
        '  sifre_hash  TEXT NOT NULL,',
        '  rol         VARCHAR(30) NOT NULL,   -- admin / doktor / hemsire / eczaci / guvenlik',
        '  aktif       BOOLEAN DEFAULT TRUE,',
        '  olusturma   TIMESTAMP DEFAULT NOW()',
        ');',
        '',
        '-- Hastalar',
        'CREATE TABLE hastalar (',
        '  id              UUID PRIMARY KEY DEFAULT gen_random_uuid(),',
        '  ad_soyad        VARCHAR(120) NOT NULL,',
        '  tc_hash         TEXT NOT NULL,       -- bcrypt hash',
        '  dogum_tarihi    DATE NOT NULL,',
        '  kan_grubu       VARCHAR(5),',
        '  alerjiler       TEXT[] DEFAULT \'{}\',',
        '  kronik          TEXT[] DEFAULT \'{}\',',
        '  doktor_id       UUID REFERENCES kullanicilar(id),',
        '  kogus_oda       VARCHAR(20) NOT NULL,',
        '  nfc_uid         VARCHAR(32) UNIQUE NOT NULL,',
        '  durum           VARCHAR(20) DEFAULT \'Stabil\',',
        '  yatis_tarihi    TIMESTAMP DEFAULT NOW(),',
        '  taburcu_tarihi  TIMESTAMP',
        ');',
        '',
        '-- İlaçlar (master katalog)',
        'CREATE TABLE ilaclar (',
        '  id          UUID PRIMARY KEY DEFAULT gen_random_uuid(),',
        '  rfid_uid    VARCHAR(32) UNIQUE NOT NULL,',
        '  adi         VARCHAR(200) NOT NULL,',
        '  etken_madde TEXT[] NOT NULL,',
        '  atc_kodu    VARCHAR(10),',
        '  lot_no      VARCHAR(50),',
        '  skt         DATE NOT NULL,',
        '  stok        INTEGER DEFAULT 0,',
        '  stok_esik   INTEGER DEFAULT 10',
        ');',
        '',
        '-- Hasta-İlaç Atamaları',
        'CREATE TABLE hasta_ilac (',
        '  id          UUID PRIMARY KEY DEFAULT gen_random_uuid(),',
        '  hasta_id    UUID REFERENCES hastalar(id) ON DELETE CASCADE,',
        '  ilac_id     UUID REFERENCES ilaclar(id),',
        '  doz_sikligi VARCHAR(30) NOT NULL,   -- "3x1", "1x2" vb.',
        '  baslangic   DATE NOT NULL,',
        '  bitis       DATE,',
        '  doktor_id   UUID REFERENCES kullanicilar(id),',
        '  aktif       BOOLEAN DEFAULT TRUE',
        ');',
        '',
        '-- İlaç Uygulama Log (her doğrulama buraya yazılır)',
        'CREATE TABLE ilac_log (',
        '  id              UUID PRIMARY KEY DEFAULT gen_random_uuid(),',
        '  hasta_id        UUID REFERENCES hastalar(id),',
        '  ilac_id         UUID REFERENCES ilaclar(id),',
        '  hemsire_id      UUID REFERENCES kullanicilar(id),',
        '  zaman           TIMESTAMP DEFAULT NOW(),',
        '  onaylandi       BOOLEAN NOT NULL,',
        '  red_nedeni      TEXT,',
        '  override_doktor UUID REFERENCES kullanicilar(id)',
        ');',
        '',
        '-- NFC/RFID Okuma Log (audit — değiştirilemez)',
        'CREATE TABLE audit_log (',
        '  id          BIGSERIAL PRIMARY KEY,',
        '  olay_tipi   VARCHAR(50) NOT NULL,  -- NFC_GIRIS / RFID_ILAC / KONUM_DEGISIM / ALARM',
        '  kaynak_uid  VARCHAR(32),',
        '  kullanici_id UUID,',
        '  hasta_id    UUID,',
        '  meta        JSONB,',
        '  zaman       TIMESTAMP DEFAULT NOW()',
        ');',
        '',
        '-- Konum (gerçek zamanlı, Redis\'ten sync)',
        'CREATE TABLE konum_gecmis (',
        '  id          BIGSERIAL PRIMARY KEY,',
        '  hasta_id    UUID REFERENCES hastalar(id),',
        '  anten_id    VARCHAR(30) NOT NULL,',
        '  bina_kat    VARCHAR(20),',
        '  oda_kodu    VARCHAR(20),',
        '  rssi        SMALLINT,',
        '  zaman       TIMESTAMP DEFAULT NOW()',
        ');',
      ]),
      spacer(),

      h2('3.3  API Endpoint Kataloğu'),
      h3('3.3.1  Kimlik Doğrulama'),
      table(
        ['Method', 'Endpoint', 'Yetki', 'Açıklama'],
        [
          ['POST', '/api/auth/login', 'Herkese Açık', 'Email + şifre → JWT access (15dk) + refresh (7gün) token'],
          ['POST', '/api/auth/refresh', 'Herkese Açık', 'Refresh token → yeni access token'],
          ['POST', '/api/auth/logout', 'Giriş yapmış', 'Refresh token geçersiz kıl (Redis blacklist)'],
          ['GET', '/api/auth/me', 'Giriş yapmış', 'Aktif kullanıcı bilgisi'],
        ],
        [900, 2400, 1600, 4460]
      ),
      spacer(),

      h3('3.3.2  Hasta Yönetimi'),
      table(
        ['Method', 'Endpoint', 'Yetki', 'Açıklama'],
        [
          ['GET', '/api/hastalar', 'Doktor / Hemşire', 'Tüm aktif hastalar; filtreleme: ?kogus=A&durum=Kritik'],
          ['POST', '/api/hastalar', 'Doktor', 'Yeni hasta kaydı; body: hasta nesnesi (şema §3.2)'],
          ['GET', '/api/hastalar/:id', 'Doktor / Hemşire', 'Tek hasta detayı + ilaç listesi + son NFC logları'],
          ['PATCH', '/api/hastalar/:id', 'Doktor', 'Kısmi güncelleme (durum, oda, doktor değişikliği)'],
          ['DELETE', '/api/hastalar/:id', 'Admin', 'Soft delete (taburcu_tarihi set edilir)'],
          ['POST', '/api/hastalar/nfc-ara', 'Hemşire+', 'Body: {uid} → Hasta profili döndürür (NFC okuma)'],
        ],
        [900, 2400, 1600, 4460]
      ),
      spacer(),

      h3('3.3.3  İlaç Doğrulama (Kritik Endpoint)'),
      table(
        ['Method', 'Endpoint', 'Yetki', 'Açıklama'],
        [
          ['POST', '/api/ilac/dogrula', 'Hemşire', '5-adım doğrulama motoru; request/response §3.4'],
          ['POST', '/api/ilac/override', 'Doktor', 'Başarısız doğrulamayı doktor onayıyla geçersiz kıl'],
          ['GET', '/api/ilac/program/:hasta_id', 'Hemşire+', 'Hastanın günlük ilaç programı + verilme durumu'],
          ['GET', '/api/ilaclar', 'Eczacı+', 'İlaç kataloğu; filtreleme: ?stok_alti=true'],
          ['POST', '/api/ilaclar', 'Eczacı', 'Yeni ilaç kaydı + RFID UID atama'],
          ['PATCH', '/api/ilaclar/:id/stok', 'Eczacı', 'Stok güncelleme'],
        ],
        [900, 2600, 1400, 4460]
      ),
      spacer(),

      h3('3.3.4  Konum & Alarm'),
      table(
        ['Method', 'Endpoint', 'Yetki', 'Açıklama'],
        [
          ['GET', '/api/konum/canli', 'Hemşire+', 'Tüm hastaların anlık konumu (Redis\'ten)'],
          ['GET', '/api/konum/:hasta_id/gecmis', 'Doktor+', 'Son 24 saatlik konum geçmişi'],
          ['POST', '/api/konum/anten-okuma', 'Sistem (API Key)', 'RFID anten okuyucularının veri gönderme endpoint\'i'],
          ['GET', '/api/alarmlar', 'Hemşire+', 'Aktif alarmlar; filtreleme: ?seviye=KRİTİK'],
          ['PATCH', '/api/alarmlar/:id/onayla', 'Hemşire+', 'Alarmı onaylandı olarak işaretle'],
          ['GET', '/api/raporlar/ilac-log', 'Doktor+', 'İlaç uygulama raporu; ?baslangic=&bitis=&hasta_id='],
        ],
        [900, 2800, 1400, 4260]
      ),
      spacer(),

      h2('3.4  İlaç Doğrulama API — Request / Response Detayı'),
      h4('Request (POST /api/ilac/dogrula)'),
      codeBlock([
        '{',
        '  "hasta_id": "uuid-xxxx",',
        '  "rfid_uid": "RF:M1:4A:2B",',
        '  "hemsire_id": "uuid-yyyy",',
        '  "okuyucu_id": "READER-A201-01"',
        '}',
      ]),
      spacer(),
      h4('Response — Başarılı'),
      codeBlock([
        '{',
        '  "sonuc": "ONAYLANDI",',
        '  "ilac": {',
        '    "id": "uuid-zzzz",',
        '    "adi": "Metoprolol 50mg",',
        '    "doz": "1x1"',
        '  },',
        '  "kontroller": {',
        '    "hasta_ilac_eslesme": true,',
        '    "doz_uygunlugu": true,',
        '    "alerji_cakismasi": false,',
        '    "skt_gecerli": true,',
        '    "cift_doz": false',
        '  },',
        '  "log_id": "bigint-nnnn",',
        '  "zaman": "2025-06-15T09:14:32Z"',
        '}',
      ]),
      spacer(),
      h4('Response — Başarısız (Alerji Çakışması)'),
      codeBlock([
        '{',
        '  "sonuc": "REDDEDİLDİ",',
        '  "hata_kodu": "ALERJI_CATISMASI",',
        '  "mesaj": "Hasta Penisilin alerjisi var. Bu ilaç Amoksisilin içermektedir.",',
        '  "etkilenen_maddeler": ["Amoksisilin trihidrat"],',
        '  "hasta_alerjileri": ["Penisilin"],',
        '  "kontroller": {',
        '    "hasta_ilac_eslesme": true,',
        '    "doz_uygunlugu": true,',
        '    "alerji_cakismasi": true,   // true = çakışma VAR = red',
        '    "skt_gecerli": true,',
        '    "cift_doz": false',
        '  },',
        '  "override_gerekli": true,',
        '  "override_endpoint": "/api/ilac/override"',
        '}',
      ]),
      pageBreak(),

      // ════════════════════════════════════════
      // 4. FRONTEND MİMARİSİ
      // ════════════════════════════════════════
      h1('4. Frontend Mimarisi'),

      h2('4.1  Sayfa & Bileşen Yapısı'),
      codeBlock([
        'src/',
        '├── pages/',
        '│   ├── Login.jsx                 // JWT login formu',
        '│   ├── Dashboard.jsx             // Ana özet ekranı (istatistikler)',
        '│   ├── HastalarPage.jsx          // Hasta listesi + detay yan panel',
        '│   ├── IlacDogrulamaPage.jsx     // Hemşire ilaç tarama ekranı',
        '│   ├── KonumHaritaPage.jsx       // Gerçek zamanlı kat planı',
        '│   ├── AlarmlarPage.jsx          // Alarm listesi + onaylama',
        '│   ├── RaporlarPage.jsx          // Rapor oluşturma + dışa aktarma',
        '│   └── AdminPage.jsx             // Kullanıcı, oda, anten yönetimi',
        '├── components/',
        '│   ├── layout/',
        '│   │   ├── Sidebar.jsx           // Sol navigasyon (rol bazlı menü)',
        '│   │   ├── Topbar.jsx            // Alarm zili + kullanıcı avatar',
        '│   │   └── ProtectedRoute.jsx    // JWT + rol kontrolü',
        '│   ├── hasta/',
        '│   │   ├── HastaKartı.jsx        // Liste öğesi bileşeni',
        '│   │   ├── HastaDetay.jsx        // Yan panel — tam profil',
        '│   │   └── NfcOkuyucu.jsx        // NFC tarama arayüzü bileşeni',
        '│   ├── ilac/',
        '│   │   ├── IlacKartı.jsx         // RFID ilaç seçimi',
        '│   │   ├── DogrulamaMotoru.jsx   // 5 adım animasyonlu kontrol',
        '│   │   └── IlacProgramı.jsx      // Günlük ilaç çizelgesi',
        '│   ├── konum/',
        '│   │   ├── KatPlani.jsx          // SVG kat planı bileşeni',
        '│   │   └── HastaNoktası.jsx      // Harita üzerinde hasta işareti',
        '│   └── ortak/',
        '│       ├── AlarmBanner.jsx       // Kritik alarm popup',
        '│       ├── StatusBadge.jsx       // Stabil/İzleme/Kritik rozetleri',
        '│       └── DataTable.jsx         // Sıralama/filtreleme destekli tablo',
        '├── store/',
        '│   ├── authStore.js             // JWT, kullanıcı bilgisi',
        '│   ├── hastaStore.js            // Hasta listesi + seçili hasta',
        '│   ├── alarmStore.js            // Aktif alarmlar',
        '│   └── konumStore.js            // Gerçek zamanlı konum verisi',
        '├── hooks/',
        '│   ├── useNfc.js               // NFC WebSerial okuyucu hook\'u',
        '│   ├── useSocket.js             // Socket.io bağlantı yönetimi',
        '│   └── useAlarm.js              // Alarm ses + bildirim hook\'u',
        '└── api/',
        '    └── client.js                // Axios instance + interceptor',
      ]),
      spacer(),

      h2('4.2  Gerçek Zamanlı WebSocket Olayları'),
      table(
        ['Olay (Event)', 'Yön', 'Payload', 'Tetikleyen'],
        [
          ['nfc:okuma', 'Server→Client', '{uid, hasta, oda, zaman}', 'NFC okuyucu UID gönderir'],
          ['rfid:ilac_tarama', 'Server→Client', '{rfid_uid, ilac, hemsire}', 'İlaç RFID okuyucuya yaklaşır'],
          ['konum:guncelleme', 'Server→Client', '{hasta_id, oda, zaman}', 'Anten okuyucu veri gönderir'],
          ['alarm:yeni', 'Server→Client', '{id, seviye, mesaj, hasta}', 'Herhangi alarm tetiklenince'],
          ['alarm:onaylandi', 'Server→Client', '{id, onaylayan}', 'Alarm onaylanınca'],
          ['ilac:onaylandi', 'Server→Client', '{log_id, hasta, ilac, hemsire}', 'Başarılı ilaç doğrulaması'],
          ['ilac:reddedildi', 'Server→Client', '{hasta, ilac, hata_kodu}', 'Başarısız doğrulama'],
        ],
        [2400, 900, 3000, 3060]
      ),
      spacer(),

      h2('4.3  NFC Donanım Köprüsü (Web Tarayıcı)'),
      para('Gerçek fiziksel NFC okuyucu (örn. ACR122U) tarayıcıya bağlanmak için iki seçenek sunulur:'),
      h4('Seçenek A: WebSerial API (Chrome 89+)'),
      codeBlock([
        '// hooks/useNfc.js',
        'export function useNfc(onRead) {',
        '  const connect = async () => {',
        '    const port = await navigator.serial.requestPort();',
        '    await port.open({ baudRate: 9600 });',
        '    const reader = port.readable.getReader();',
        '    while (true) {',
        '      const { value, done } = await reader.read();',
        '      if (done) break;',
        '      const uid = parseUID(value);  // Donanıma göre parse',
        '      onRead(uid);',
        '    }',
        '  };',
        '  return { connect };',
        '}',
      ]),
      spacer(),
      h4('Seçenek B: Yerel Agent (Node.js Desktop App)'),
      para('NFC okuyucu, kullanıcının bilgisayarında çalışan küçük bir Node.js ajanı tarafından okunur. Ajan okunan UID\'i WebSocket üzerinden backend\'e iletir. Bu seçenek tüm tarayıcılarda çalışır.'),
      pageBreak(),

      // ════════════════════════════════════════
      // 5. GÜVENLİK
      // ════════════════════════════════════════
      h1('5. Güvenlik Gereksinimleri'),

      h2('5.1  Kimlik Doğrulama & Yetkilendirme'),
      bullet('JWT access token süresi: 15 dakika. Refresh token süresi: 7 gün (Redis\'te).'),
      bullet('Tüm şifreler bcrypt ile hash\'lenir (cost factor: 12).'),
      bullet('Her endpoint için rol bazlı middleware çalışır (bkz. §3.3).'),
      bullet('Ardışık 5 başarısız giriş denemesinde hesap 30 dakika kilitlenir (Redis sayacı).'),
      bullet('TC kimlik numaraları veritabanında asla düz metin saklanmaz; bcrypt hash\'i + ayrı tuzlama.'),
      spacer(),

      h2('5.2  Veri Güvenliği'),
      bullet('HTTPS zorunludur. HTTP istekleri 301 yönlendirmesiyle HTTPS\'e çevrilir.'),
      bullet('Veritabanı bağlantısı TLS ile şifreli.'),
      bullet('RFID/NFC UID\'leri log tablosunda AES-256-GCM ile şifreli saklanır.'),
      bullet('SQL enjeksiyonuna karşı Prisma ORM\'nin parametreli sorguları kullanılır.'),
      bullet('XSS\'e karşı tüm kullanıcı girdileri DOMPurify ile sanitize edilir.'),
      bullet('CSRF tokenları form gönderimlerinde zorunludur.'),
      spacer(),

      h2('5.3  Audit & Uyumluluk'),
      bullet('Her veri değişikliği (CREATE / UPDATE / DELETE) audit_log tablosuna yazılır.'),
      bullet('Audit log kayıtları silinemez ve güncellenemez (PostgreSQL row security policy).'),
      bullet('İlaç uygulama kayıtları 10 yıl süreyle saklanır (Türkiye sağlık mevzuatı gerekliliği).'),
      bullet('Kişisel sağlık verileri KVKK ve GDPR gerekliliklerine uygun işlenir.'),
      pageBreak(),

      // ════════════════════════════════════════
      // 6. PERFORMANS GEREKSİNİMLERİ
      // ════════════════════════════════════════
      h1('6. Performans & Ölçeklenebilirlik'),

      h2('6.1  Hedef Metrikler'),
      table(
        ['Metrik', 'Hedef Değer', 'Ölçüm Yöntemi'],
        [
          ['API yanıt süresi (p95)', '< 200 ms', 'k6 yük testi'],
          ['İlaç doğrulama yanıt süresi', '< 500 ms', 'End-to-end test'],
          ['WebSocket mesaj gecikmesi', '< 100 ms', 'Socket latency log'],
          ['Eşzamanlı kullanıcı kapasitesi', '500 kullanıcı', 'k6 + Node.js cluster'],
          ['NFC okuma → ekran güncelleme', '< 1 saniye', 'Manuel + Playwright test'],
          ['Sistem çalışma süresi (uptime)', '% 99.5', 'UptimeRobot izleme'],
          ['Sayfa ilk yükleme (LCP)', '< 2.5 saniye', 'Lighthouse'],
        ],
        [3000, 2200, 4160]
      ),
      spacer(),

      h2('6.2  Önbellekleme Stratejisi'),
      bullet('Aktif hasta listesi: Redis\'te 60 saniye TTL (sık okunur, nadiren değişir).'),
      bullet('İlaç kataloğu: Redis\'te 10 dakika TTL.'),
      bullet('Gerçek zamanlı konum: Redis hash\'i (hasta_id → konum); TTL yok, her anten okumasında güncellenir.'),
      bullet('JWT blacklist (logout edilen tokenlar): Redis\'te token süresi kadar TTL.'),
      pageBreak(),

      // ════════════════════════════════════════
      // 7. TEST STRATEJİSİ
      // ════════════════════════════════════════
      h1('7. Test Stratejisi'),

      h2('7.1  Birim Testler'),
      para('İlaç doğrulama motorunun 5 adımı, tüm edge case\'lerle birim test kapsamına alınmalıdır:'),
      table(
        ['Test Senaryosu', 'Beklenen Sonuç', 'Adım'],
        [
          ['Hasta listesinde olmayan ilaç taranır', 'REDDEDİLDİ — HASTA_ILAC_ESLESME_YOK', '1'],
          ['Günlük maksimum doz aşılmış', 'REDDEDİLDİ — MAX_DOZ_ASILDI', '2'],
          ['Penisilin alerjisi olan hastaya Amoksisilin', 'REDDEDİLDİ — ALERJI_CATISMASI', '3'],
          ['SKT\'si 2024-01 olan ilaç (geçmiş)', 'REDDEDİLDİ — SKT_GECMIS', '4'],
          ['Aynı ilaç 3 saat önce verilmiş (4 saatlik pencere)', 'REDDEDİLDİ — CIFT_DOZ', '5'],
          ['Tüm kontroller geçilir', 'ONAYLANDI', 'Tümü'],
          ['Doktor override ile alerji çakışması geçilir', 'ONAYLANDI (override_log kaydı)', '3'],
        ],
        [3400, 3000, 2960]
      ),
      spacer(),

      h2('7.2  Entegrasyon Testleri'),
      bullet('NFC okuma → hasta profili yüklenme → audit_log kaydı akışı.'),
      bullet('RFID tarama → 5 adım motor → ilaç_log kaydı → WebSocket bildirimi akışı.'),
      bullet('Anten okuma → Redis konum güncelleme → WebSocket broadcast → kat planı animasyonu.'),
      bullet('İzin dışı bölge tespiti → alarm_log kaydı → bildirim gönderimi.'),
      spacer(),

      h2('7.3  Uçtan Uca (E2E) Testler'),
      para('Playwright kullanılarak aşağıdaki kullanıcı senaryoları otomatize edilir:'),
      numbered('Hemşire giriş yapar → hasta seçer → NFC okur → hasta profili açılır → çıkış.'),
      numbered('Hemşire ilaç RFID tarar → alerji çakışması → sistem reddeder → doktor override → onay.'),
      numbered('Güvenlik görevlisi → konum haritasını açar → izin dışı alarm görür → onaylar.'),
      pageBreak(),

      // ════════════════════════════════════════
      // 8. KURULUM & DAĞITIM
      // ════════════════════════════════════════
      h1('8. Kurulum & Dağıtım'),

      h2('8.1  Geliştirme Ortamı Kurulumu'),
      codeBlock([
        '# 1. Repository klonla',
        'git clone https://github.com/org/medtrack.git',
        'cd medtrack',
        '',
        '# 2. Bağımlılıkları yükle',
        'cd backend && npm install',
        'cd ../frontend && npm install',
        '',
        '# 3. Ortam değişkenlerini ayarla',
        'cp backend/.env.example backend/.env',
        '# .env dosyasını düzenle (bkz. §8.2)',
        '',
        '# 4. Docker ile PostgreSQL ve Redis başlat',
        'docker compose up -d db redis',
        '',
        '# 5. Veritabanı migration ve seed',
        'cd backend && npx prisma migrate dev',
        'npx prisma db seed',
        '',
        '# 6. Geliştirme sunucularını başlat',
        'npm run dev          # backend :3001',
        'cd ../frontend && npm run dev   # frontend :5173',
      ]),
      spacer(),

      h2('8.2  Ortam Değişkenleri (.env)'),
      codeBlock([
        '# Veritabanı',
        'DATABASE_URL="postgresql://medtrack:gizli@localhost:5432/medtrack"',
        '',
        '# Redis',
        'REDIS_URL="redis://localhost:6379"',
        '',
        '# JWT',
        'JWT_SECRET="en-az-64-karakter-rastgele-gizli-anahtar"',
        'JWT_REFRESH_SECRET="baska-bir-64-karakter-gizli-anahtar"',
        '',
        '# Şifreleme (RFID UID\'leri için)',
        'ENCRYPTION_KEY="32-byte-hex-aes-256-gcm-anahtari"',
        '',
        '# Sunucu',
        'PORT=3001',
        'FRONTEND_URL="https://medtrack.hastane.com"',
        '',
        '# E-posta (alarm bildirimleri)',
        'SMTP_HOST="smtp.hastane.com"',
        'SMTP_PORT=587',
        'SMTP_USER="medtrack@hastane.com"',
        'SMTP_PASS="smtp-sifresi"',
        '',
        '# Opsiyonel: SMS (Twilio)',
        'TWILIO_SID="ACxxxx"',
        'TWILIO_TOKEN="xxxx"',
        'TWILIO_FROM="+905xxxxxxxxx"',
      ]),
      spacer(),

      h2('8.3  Üretim Dağıtımı (Docker Compose)'),
      codeBlock([
        'services:',
        '  backend:',
        '    build: ./backend',
        '    environment:',
        '      - NODE_ENV=production',
        '    depends_on: [db, redis]',
        '    restart: unless-stopped',
        '',
        '  frontend:',
        '    build: ./frontend',
        '    environment:',
        '      - VITE_API_URL=https://api.medtrack.hastane.com',
        '',
        '  nginx:',
        '    image: nginx:1.25-alpine',
        '    volumes:',
        '      - ./nginx.conf:/etc/nginx/nginx.conf',
        '      - ./ssl:/etc/ssl/medtrack',
        '    ports: ["80:80", "443:443"]',
        '',
        '  db:',
        '    image: postgres:16-alpine',
        '    volumes: [pg_data:/var/lib/postgresql/data]',
        '',
        '  redis:',
        '    image: redis:7-alpine',
        '    command: redis-server --requirepass gizli',
        '',
        'volumes:',
        '  pg_data:',
      ]),
      pageBreak(),

      // ════════════════════════════════════════
      // 9. GELİŞTİRME YOLU HARİTASI
      // ════════════════════════════════════════
      h1('9. Geliştirme Yol Haritası'),

      h2('9.1  Faz 1 — Temel Sistem (4 Hafta)'),
      table(
        ['Hafta', 'Görevler', 'Çıktı'],
        [
          ['1', 'Proje iskeleti (React + Express + PostgreSQL + Docker), kimlik doğrulama (JWT, roller), veritabanı şeması + migration', 'Çalışan login ekranı, rol bazlı yönlendirme'],
          ['2', 'Hasta CRUD API + listesi + detay sayfası, NFC simülasyon modu (gerçek donanım olmadan), audit_log altyapısı', 'Hasta yönetimi tam çalışır'],
          ['3', 'İlaç kataloğu CRUD, 5-adım doğrulama motoru (backend), doğrulama UI bileşeni (animasyonlu), ilaç_log tablosu', 'İlaç doğrulama çalışır'],
          ['4', 'WebSocket kurulumu, gerçek zamanlı alarm sistemi, konum haritası (SVG kat planı), alarm onaylama akışı', 'Faz 1 tamamlandı — demo hazır'],
        ],
        [900, 5800, 2660]
      ),
      spacer(),

      h2('9.2  Faz 2 — Gelişmiş Özellikler (3 Hafta)'),
      table(
        ['Hafta', 'Görevler', 'Çıktı'],
        [
          ['5', 'Raporlama modülü (PDF/Excel dışa aktarım), ilaç programı görünümü, gelişmiş filtreleme/arama', 'Raporlama çalışır'],
          ['6', 'Fiziksel NFC okuyucu entegrasyonu (ACR122U + WebSerial), RFID anten simülasyonu → gerçek anten API köprüsü', 'Gerçek donanım desteği'],
          ['7', 'Performans testleri (k6), güvenlik denetimi (OWASP Top 10), E2E testler (Playwright), dokümantasyon', 'Üretim hazır'],
        ],
        [900, 5800, 2660]
      ),
      spacer(),

      h2('9.3  Faz 3 — Opsiyonel İyileştirmeler'),
      bullet('Mobil uygulama (React Native) — hemşirelerin tablet/telefon üzerinden kullanımı.'),
      bullet('AI alerji tahmini — literatür veritabanıyla etken madde çakışma tahmini.'),
      bullet('HL7 FHIR entegrasyonu — hastane HIS/EMR sistemleriyle veri senkronizasyonu.'),
      bullet('Biyometrik doğrulama — parmak izi veya yüz tanıma ile hemşire kimliği.'),
      pageBreak(),

      // ════════════════════════════════════════
      // 10. SÖZLÜK
      // ════════════════════════════════════════
      h1('10. Terimler Sözlüğü'),
      table(
        ['Terim', 'Açıklama'],
        [
          ['NFC', 'Near Field Communication — 13.56 MHz\'de çalışan, ≤10 cm mesafeli kablosuz iletişim standardı'],
          ['RFID', 'Radio Frequency Identification — radyo dalgalarıyla etiket okuma teknolojisi'],
          ['UID', 'Unique Identifier — NFC/RFID etiketinin fabrikadan gelen değiştirilemez benzersiz kimliği'],
          ['HF RFID', 'High Frequency RFID — 13.56 MHz, ≤10 cm kapsama, ilaç etiketleri için kullanılır'],
          ['UHF RFID', 'Ultra High Frequency RFID — 860–960 MHz, 1–10 m kapsama, konum takibi için kullanılır'],
          ['RSSI', 'Received Signal Strength Indicator — sinyal gücü değeri; konum belirleme algoritmasında kullanılır'],
          ['JWT', 'JSON Web Token — durumsuz kimlik doğrulama tokeni'],
          ['ATC', 'Anatomical Therapeutic Chemical — WHO\'nun uluslararası ilaç sınıflandırma sistemi'],
          ['SKT', 'Son Kullanma Tarihi — ilacın güvenli kullanım son tarihi'],
          ['Override', 'Yetkili kullanıcının (doktor) sistem kısıtlamasını geçersiz kılması; her zaman loglanır'],
          ['WebSerial', 'Tarayıcının USB seri port cihazlarla (NFC okuyucu) iletişimini sağlayan W3C API\'si'],
          ['Broadcast', 'WebSocket üzerinden tüm bağlı istemcilere eşzamanlı mesaj gönderimi'],
          ['KVKK', 'Kişisel Verilerin Korunması Kanunu — Türkiye'],
          ['FHIR', 'Fast Healthcare Interoperability Resources — sağlık verisi değişim standardı (HL7)'],
        ],
        [2400, 6960]
      ),
      spacer(2),
      infoBox('Bu doküman hakkında sorularınız veya geliştirme sürecinde ihtiyaç duyduğunuz ek detaylar için proje sahibiyle iletişime geçiniz. Doküman versiyon kontrolü Git üzerinden yönetilmelidir.', GREEN_BG, GREEN),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('/mnt/user-data/outputs/MedTrack_Teknik_Dokumantasyon.docx', buf);
  console.log('OK');
});
