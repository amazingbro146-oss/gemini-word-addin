Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    const savedKey = localStorage.getItem("gemini_api_key");
    if (savedKey) {
      document.getElementById("apiKey").value = savedKey;
      document.getElementById("keyStatus").innerText = "API Key tersimpan.";
      muatDaftarModel();
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
  document.getElementById("keyStatus").innerText = "API Key tersimpan! Memuat daftar model...";
  muatDaftarModel();
}

// Mengambil daftar model aktif langsung dari Google API Key Anda
async function muatDaftarModel() {
  const apiKey = localStorage.getItem("gemini_api_key") || document.getElementById("apiKey").value.trim();
  if (!apiKey) return;

  const select = document.getElementById("modelSelect");
  const statusEl = document.getElementById("keyStatus");

  try {
    const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey}`);
    const data = await res.json();

    if (data.models && data.models.length > 0) {
      const validModels = data.models
        .filter(m => m.supportedGenerationMethods && m.supportedGenerationMethods.includes("generateContent") && m.name.includes("gemini"))
        .map(m => m.name.replace("models/", ""));

      if (validModels.length > 0) {
        select.innerHTML = "";
        // Urutkan agar model flash / lite berada di atas (lebih jarang sibuk)
        validModels.sort((a, b) => {
          if (a.includes("lite") && !b.includes("lite")) return -1;
          if (!a.includes("lite") && b.includes("lite")) return 1;
          if (a.includes("flash") && !b.includes("flash")) return -1;
          if (!a.includes("flash") && b.includes("flash")) return 1;
          return b.localeCompare(a);
        });

        validModels.forEach(modelName => {
          const opt = document.createElement("option");
          opt.value = modelName;
          opt.text = modelName;
          select.appendChild(opt);
        });
        statusEl.innerText = `Berhasil memuat ${validModels.length} model aktif.`;
      }
    } else if (data.error) {
      statusEl.innerText = "Gagal memuat model: " + data.error.message;
    }
  } catch (e) {
    statusEl.innerText = "Gagal mengecek daftar model.";
  }
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
  const select = document.getElementById("modelSelect");
  const selectedModel = select.value;
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

  // Ambil semua model yang ada di dropdown sebagai antrean cadangan
  let antreanModel = [selectedModel];
  if (useFallback) {
    for (let i = 0; i < select.options.length; i++) {
      const m = select.options[i].value;
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
        return;
      } else if (data.error) {
        const pesanError = data.error.message || "";
        const isSibukAtauTidakTersedia =
          pesanError.toLowerCase().includes("high demand") ||
          pesanError.toLowerCase().includes("no longer available") ||
          pesanError.toLowerCase().includes("not found") ||
          data.error.code === 503 ||
          data.error.code === 429 ||
          data.error.code === 404;

        if (isSibukAtauTidakTersedia && i < antreanModel.length - 1) {
          outputBox.innerText = `${modelSaatIni} sedang penuh/tidak tersedia, mencoba ${antreanModel[i + 1]}...`;
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
  if (!hasilTeks || hasilTeks === "Jawaban akan muncul di sini..." || hasilTeks.startsWith("Sedang memproses") || hasilTeks.includes("mencoba ")) {
    return;
  }

  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.insertText(hasilTeks, Word.InsertLocation.replace);
    await context.sync();
  });
}
