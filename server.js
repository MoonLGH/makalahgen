const express = require("express");
const fs = require("fs");
const path = require("path");
const {imageSize} = require("image-size"); // Tambahkan di baris paling atas
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const multer = require("multer"); // Tambahan untuk upload
const ImageModule = require("docxtemplater-image-module-free"); // Tambahan untuk image di docx

const app = express();
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// Konfigurasi Penyimpanan File Upload Sementara
const upload = multer({ dest: "uploads/" });

// --- HELPER FUNCTIONS ---
app.use("/image", express.static(path.join(__dirname, "image")));

// Fungsi baru untuk menangani indentasi secara dinamis
const formatTextWithIndent = (text, indentType) => {
  if (!text) return [];

  // Filter untuk menghapus baris kosong yang tidak sengaja terbuat
  const lines = text.split(/\r?\n/).filter(line => line.trim() !== "");

  return lines.map((line, index) => {
    let prefix = "";
    
    if (indentType === "all") {
      prefix = "\t"; // Semua baris menjorok
    } else if (indentType === "first") {
      // Hanya indeks 0 (baris pertama dari blok teks tersebut) yang menjorok
      prefix = (index === 0) ? "\t" : "";
    } else if (indentType === "none") {
      prefix = ""; // Tidak ada yang menjorok
    } else {
      prefix = "\t"; // Fallback
    }

    return {
      text: prefix + line
    };
  });
};

const formatReferensiOtomatis = (data, style) => {
  const refs = [];
  Object.keys(data).forEach(key => {
    if (key.startsWith("ref_penulis_")) {
      const id = key.split("_")[2];
      const p = data[`ref_penulis_${id}`]?.trim();
      const t = data[`ref_tahun_${id}`]?.trim();
      const j = data[`ref_judul_${id}`]?.trim();
      const b = data[`ref_penerbit_${id}`]?.trim();
      const k = data[`ref_kota_${id}`]?.trim();

      if (p && j) {
        let formatted = "";
        if (style === "APA") {
          formatted = `${p}. (${t}). ${j}. ${k ? k + ": " : ""}${b}.`;
        } else if (style === "MLA") {
          formatted = `${p}. "${j}." ${b}, ${t}.`;
        } else {
          formatted = `${p}. ${j}. ${k ? k + ": " : ""}${b}, ${t}.`;
        }
        refs.push(formatted);
      }
    }
  });

  refs.sort((a, b) => a.localeCompare(b));
  return refs.map(item => ({ text: item }));
};

// --- ROUTES ---

app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "index.html"));
});

// Update route /generate untuk menerima single file 'logo_custom'
app.post("/generate", upload.single("logo_custom"), (req, res) => {
  try {
    const templatePath = path.resolve(__dirname, "template.docx");
    if (!fs.existsSync(templatePath)) return res.status(404).send("File template.docx tidak ditemukan");

    const content = fs.readFileSync(templatePath, "binary");
    const zip = new PizZip(content);

const imageOptions = {
      centered: true,
      getImage: (tagValue) => fs.readFileSync(tagValue),
      getSize: (imgBuffer) => {
        try {
          const dimensions = imageSize(imgBuffer);
          
          // Tentukan lebar target yang ideal untuk cover makalah (dalam pixel)
          // 150-180px biasanya pas untuk logo yang lebar seperti UNNES/UNDIP
          const targetWidth = 300; 

          // Rumus Rasio: Lebar Asli / Tinggi Asli
          const originalRatio = dimensions.width / dimensions.height;

          // Hitung tinggi berdasarkan rasio asli supaya TIDAK GEPENG
          // Tinggi = Lebar Target / Rasio
          const targetHeight = targetWidth / originalRatio;

          // Log untuk debugging di terminal (cek apakah rasionya benar)
          console.log(`Original: ${dimensions.width}x${dimensions.height} (Ratio: ${originalRatio})`);
          console.log(`Target: ${targetWidth}x${targetHeight}`);

          return [Math.round(targetWidth), Math.round(targetHeight)];
        } catch (e) {
          console.error("Gagal memproses dimensi gambar:", e);
          return [300, 300]; // Fallback jika terjadi error pembacaan
        }
      },
    };
    const imageModule = new ImageModule(imageOptions);
    const doc = new Docxtemplater(zip, {
      delimiters: { start: "[[", end: "]]" },
      paragraphLoop: true,
      linebreaks: true,
      modules: [imageModule], // Daftarkan module image
    });

    // Tentukan path logo: Pakai yang diupload, jika tidak ada pakai default UNNES
// Tentukan path logo: 
    // Jika ada file yang diupload, pakai path file tersebut.
    // Jika tidak ada (null/undefined), pakai logo UNNES sebagai default.
    const logoPath = req.file 
      ? req.file.path 
      : path.resolve(__dirname, "image/unnes.png"); 
      // ^ Pastikan nama file di folder 'image' adalah unnes.png atau sesuaikan namanya
    // Proses Anggota
    const anggotaArray = [];
    Object.keys(req.body).forEach(key => {
      if (key.startsWith("anggota_nama_")) {
        const index = key.split("_")[2];
        anggotaArray.push({
          nama: req.body[key],
          nim: req.body["anggota_nim_" + index]
        });
      }
    });

    // Proses BAB & Sub-BAB beserta gaya indentasinya
    const babArray = [];
    Object.keys(req.body).forEach(key => {
      if (key.startsWith("bab_judul_")) {
        const index = key.split("_")[2];
        const subBabArray = [];
        
        Object.keys(req.body).forEach(subKey => {
          if (subKey.startsWith("subbab_judul_" + index + "_")) {
            const subIndex = subKey.split("_")[3];
            
            subBabArray.push({
              subjudul: req.body[subKey],
              IsiSubBab: formatTextWithIndent(
                req.body["subbab_isi_" + index + "_" + subIndex],
                req.body["indent_subbab_" + index + "_" + subIndex]
              )
            });
          }
        });
        
        babArray.push({
          nomor: req.body["bab_nomor_" + index],
          judul: req.body[key],
          IsiBab: formatTextWithIndent(
            req.body["bab_isi_" + index], 
            req.body["indent_bab_" + index]
          ),
          SubBab: subBabArray
        });
      }
    });

    const formattedReferensi = formatReferensiOtomatis(req.body, req.body.referensi_style);

    doc.setData({
      ...req.body,
      logo_image: logoPath, // Tag [[%logo_image]] di Word
      "Isi-KataPengantar": formatTextWithIndent(
        req.body["Isi-KataPengantar"], 
        req.body["indent_katapengantar"]
      ),
      Anggota: anggotaArray,
      BAB: babArray,
      Referensi: formattedReferensi
    });

    doc.render();

    // Hapus file temporary upload setelah di-render ke Word
    if (req.file) {
        try { fs.unlinkSync(req.file.path); } catch(e) { console.error("Gagal hapus temp file", e); }
    }

    const buffer = doc.getZip().generate({ type: "nodebuffer", compression: "DEFLATE" });

    res.setHeader("Content-Disposition", "attachment; filename=makalah.docx");
    res.send(buffer);

  } catch (err) {
    console.error(err);
    res.status(500).send("Gagal generate: " + err.message);
  }
});

app.listen(3000, () => console.log("Server running on http://localhost:3000"));
module.exports = app;