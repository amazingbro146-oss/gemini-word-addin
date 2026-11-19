const DAFTAR_MODEL_CADANGAN = [
  "gemini-3.8-flash",
  "gemini-3.8-flash-lite",
  "gemini-3.5-flash",
  "gemini-3.8-pro",
  "gemini-3.5-pro"
];

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    const savedKey = localStorage.getItem("gemini_api_key");
    if (savedKey) {
      document.getElementById("apiKey").value = savedKey;
      document.getElementById("keyStatus").innerText = "API Key tersimpan.";
    }
  }
});

function simpanApiKey() {
  const key = document.getElementById("apiKey").value.trim();
  if (!key) {
    alert("Masukkan API Key terlebih dahulu!");
    return;
  }
  localStorage.setItem("gemini_api_key", key);
  document.getElementById("keyStatus").innerText = "API Key berhasil disimpan!";
}

async function prosesTeksTerpilih(instruksiAwal) {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load("text");
    await context.sync();

    if (!selection.text.trim()) {
      alert("Silakan blok/seleksi teks di dokumen Word terlebih dahulu!");
      return;
    }

    document.getElementById("promptInput").value = `${instruksiAwal}\n\n"${selection.text}"`;
    kirimKeGemini();
  });
}

async function panggilApiModel(apiKey, modelName, prompt) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${apiKey}`;
  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      contents: [{ parts: [{ text: prompt }] }]
    })
  });
  return await response.json();
}

async function kirimKeGemini() {
  const apiKey = localStorage.getItem("gemini_api_key") || document.getElementById("apiKey").value.trim();
  const selectedModel = document.getElementById("modelSelect").value;
  const useFallback = document.getElementById("autoFallback").checked;
  const prompt = document.getElementById("promptInput").value.trim();
  const outputBox = document.getElementById("outputBox");

  if (!apiKey) {
    alert("Harap masukkan dan simpan Gemini API Key terlebih dahulu!");
    return;
  }
  if (!prompt) {
    alert("Tulis perintah terlebih dahulu!");
    return;
  }

  // Susun urutan model yang akan dicoba (model pilihan pertama, lalu cadangan jika dicentang)
  let antreanModel = [selectedModel];
  if (useFallback) {
    for (const m of DAFTAR_MODEL_CADANGAN) {
      if (!antreanModel.includes(m)) {
        antreanModel.push(m);
      }
    }
  }

  for (let i = 0; i < antreanModel.length; i++) {
    const modelSaatIni = antreanModel[i];
    outputBox.innerText = `Sedang memproses dengan ${modelSaatIni}...`;

    try {
      const data = await panggilApiModel(apiKey, modelSaatIni, prompt);

      if (data.candidates && data.candidates[0].content) {
        outputBox.innerText = data.candidates[0].content.parts[0].text;
        return; // Berhasil, hentikan perulangan
      } else if (data.error) {
        const pesanError = data.error.message || "";
        const isHighDemand = pesanError.toLowerCase().includes("high demand") || data.error.code === 503 || data.error.code === 429;

        // Jika server sibuk dan masih ada model cadangan di antrean, lanjut ke model berikutnya
        if (isHighDemand && i < antreanModel.length - 1) {
          outputBox.innerText = `${modelSaatIni} sedang penuh, mengalihkan otomatis ke ${antreanModel[i + 1]}...`;
          continue;
        } else {
          outputBox.innerText = `Error API (${modelSaatIni}): ${pesanError}`;
          return;
        }
      } else {
        outputBox.innerText = "Tidak ada respons dari Gemini.";
        return;
      }
    } catch (err) {
      outputBox.innerText = "Gagal terhubung ke server Gemini: " + err.message;
      return;
    }
  }
}

async function sisipkanKeWord() {
  const hasilTeks = document.getElementById("outputBox").innerText;
  if (!hasilTeks || hasilTeks === "Jawaban akan muncul di sini..." || hasilTeks.startsWith("Sedang memproses") || hasilTeks.includes("mengalihkan otomatis")) {
    return;
  }

  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.insertText(hasilTeks, Word.InsertLocation.replace);
    await context.sync();
  });
}
