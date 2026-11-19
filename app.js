Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    const savedKey = localStorage.getItem("gemini_api_key");
    if (savedKey) {
      document.getElementById("apiKey").value = savedKey;
      document.getElementById("keyStatus").innerText = "API Key tersimpan di browser.";
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

// Mengambil teks yang sedang diblok di dokumen Word
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

// Mengirim prompt ke Google Gemini API (Model terbaru: gemini-3.8-flash)
async function kirimKeGemini() {
  const apiKey = localStorage.getItem("gemini_api_key") || document.getElementById("apiKey").value.trim();
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

  outputBox.innerText = "Sedang memproses jawaban...";

  try {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3.8-flash:generateContent?key=${apiKey}`;
    const response = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }]
      })
    });

    const data = await response.json();

    if (data.candidates && data.candidates[0].content) {
      outputBox.innerText = data.candidates[0].content.parts[0].text;
    } else if (data.error) {
      outputBox.innerText = "Error API: " + data.error.message;
    } else {
      outputBox.innerText = "Tidak ada respons dari Gemini.";
    }
  } catch (err) {
    outputBox.innerText = "Gagal terhubung ke server Gemini: " + err.message;
  }
}

// Menyisipkan teks jawaban ke posisi kursor di Word
async function sisipkanKeWord() {
  const hasilTeks = document.getElementById("outputBox").innerText;
  if (!hasilTeks || hasilTeks === "Jawaban akan muncul di sini..." || hasilTeks === "Sedang memproses jawaban...") {
    return;
  }

  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.insertText(hasilTeks, Word.InsertLocation.replace);
    await context.sync();
  });
}
