const axios = require("axios");
const xlsx = require("xlsx");
const fs = require("fs");
const path = require("path");

// ---------------- Configuração ----------------
const CHATPRO_BASE_URL = "https://api.chatpro.com.br";  // Base da API ChatPro
const INSTANCE_ID = "SEU_INSTANCE_ID";                // Fornecido pelo painel ChatPro
const TOKEN = "SEU_TOKEN_DE_AUTENTICACAO";            // Token de autenticação

const ARQUIVO_XLSX = "testes-numeros.xlsx";
const videosDir = path.join(__dirname, "videos");

// -------- Função para enviar texto --------
async function enviarTexto(numero, texto) {
  const url = `${CHATPRO_BASE_URL}/v1/instance/${INSTANCE_ID}/send_message`;
  
  const body = {
    number: numero,
    type: "text",
    text: { body: texto }
  };

  const headers = {
    Authorization: `Bearer ${TOKEN}`,
    "Content-Type": "application/json"
  };

  return axios.post(url, body, { headers });
}

// -------- Função para enviar vídeo --------
async function enviarVideo(numero, videoUrl, legenda) {
  const url = `${CHATPRO_BASE_URL}/v1/instance/${INSTANCE_ID}/send_message`;
  
  const body = {
    number: numero,
    type: "video",
    video: {
      link: videoUrl,
      caption: legenda
    }
  };

  const headers = {
    Authorization: `Bearer ${TOKEN}`,
    "Content-Type": "application/json"
  };

  return axios.post(url, body, { headers });
}

// ---------------- Carregar planilha ----------------
if (!fs.existsSync(ARQUIVO_XLSX)) {
  console.error(`Arquivo não encontrado: ${ARQUIVO_XLSX}`);
  process.exit(1);
}

const workbook = xlsx.readFile(ARQUIVO_XLSX);
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const contatos = xlsx.utils.sheet_to_json(sheet);

// ---------------- Loop de envios ----------------
(async () => {
  console.log(`Total de contatos: ${contatos.length}`);

  for (const contato of contatos) {
    let celular = String(contato.Coluna3 || "").replace(/\D/g, "").trim();
    if (!celular.startsWith("55")) celular = `55${celular}`;

    const mensagem = `Olá ${contato.Coluna2}, tudo bem?\n*Novidades importantes!*`;

    try {
      // Envia texto
      const textResp = await enviarTexto(celular, mensagem);
      console.log("Texto enviado:", textResp.data);

      // Envia um dos vídeos aleatórios como exemplo
      const files = fs.readdirSync(videosDir);
      const randomVideo = files[Math.floor(Math.random() * files.length)];
      const videoUrl = `https://meuservidor.com/videos/${randomVideo}`; // vídeo acessível publicamente

      const videoResp = await enviarVideo(celular, videoUrl, mensagem);
      console.log("Vídeo enviado:", videoResp.data);

    } catch (err) {
      console.error("Erro enviando para", celular, err.response?.data || err.message);
    }
  }
})();
