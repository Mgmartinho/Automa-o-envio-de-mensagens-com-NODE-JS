const xlsx = require("xlsx");
const { Client, LocalAuth, MessageMedia } = require("whatsapp-web.js");
const fs = require("fs");
const cliProgress = require("cli-progress");
const path = require("path");
const { log } = require("console");

// ---------------- Carregar os 5 vídeos aleatórios ----------------
const videos = [
  MessageMedia.fromFilePath(path.join(__dirname, "videos", "CruzAzul-1.mp4")),
  MessageMedia.fromFilePath(path.join(__dirname, "videos", "CruzAzul-2.mp4")),
  MessageMedia.fromFilePath(path.join(__dirname, "videos", "CruzAzul-3.mp4")),
  MessageMedia.fromFilePath(path.join(__dirname, "videos", "CruzAzul-4.mp4")),
  MessageMedia.fromFilePath(path.join(__dirname, "videos", "CruzAzul-5.mp4")),
];

// ---------------- Mensagens aleatórias ----------------
const mensagensAleatorias = [
  "Cruz Azul preparou novidades para você! Com melhorias e novos serviços, estamos sempre buscando oferecer o melhor.",
  "Tem novidades da Cruz Azul! Estamos trazendo melhorias e novos serviços para você.",
  "A Cruz Azul trouxe novas melhorias e serviços feitos especialmente para você.",
  "Confira as novidades que a Cruz Azul preparou! Estamos sempre inovando por você.",
  "Novos serviços e melhorias foram lançados pela Cruz Azul para melhor atender você.",
  "A Cruz Azul está trazendo novidades imperdíveis para você! Melhoria constante no atendimento.",
  "Veja as novidades da Cruz Azul! Estamos sempre buscando oferecer o melhor.",
  "Cruz Azul preparou novos serviços e melhorias pensadas especialmente para você.",
  "Novidades importantes da Cruz Azul chegaram! Sempre buscando melhorar.",
  "Cruz Azul trouxe novidades para você! Melhorias constantes no atendimento.",
  "A Cruz Azul está com novidades que preparamos especialmente para você. Seguimos trabalhando para melhorar sempre.",
  "Chegaram novas melhorias e serviços da Cruz Azul! Tudo pensado para oferecer uma experiência ainda melhor.",
  "A Cruz Azul acaba de lançar novidades importantes para você. Estamos sempre evoluindo para atender melhor.",
  "Tem coisa boa chegando! A Cruz Azul trouxe novas melhorias e serviços feitos para você.",
  "A Cruz Azul segue avançando e trouxe novidades que vão melhorar ainda mais seu atendimento.",
  "Novos serviços e melhorias já estão disponíveis na Cruz Azul, tudo planejado com cuidado para você.",
  "A Cruz Azul preparou atualizações e novidades para oferecer um atendimento cada vez melhor.",
  "Chegaram novas funcionalidades e melhorias da Cruz Azul! Tudo para beneficiar você.",
  "A Cruz Azul está cheia de novidades! Seguimos inovando e trazendo melhorias contínuas.",
  "Veja o que a Cruz Azul preparou: novas melhorias e serviços pensados para facilitar seu dia a dia.",
  "Tem atualização nova da Cruz Azul no ar! Melhorias e serviços desenvolvidos especialmente para você.",
  "A Cruz Azul continua evoluindo e trouxe novidades importantes para melhorar sua experiência.",
  "Mais novidades chegando na Cruz Azul! Estamos investindo em melhorias para atender ainda melhor.",
  "A Cruz Azul lançou novas funcionalidades e melhorias criadas com foco no seu bem-estar.",
  "Tem novidade fresquinha da Cruz Azul! Seguimos buscando sempre oferecer o melhor para você.",
  "Novos recursos e melhorias foram implementados pela Cruz Azul, tudo pensado no seu atendimento.",
  "A Cruz Azul traz mais uma rodada de melhorias e novidades feitas especialmente para você.",
  "Atualizações importantes foram realizadas pela Cruz Azul para trazer mais qualidade e conforto ao seu atendimento.",
  "A Cruz Azul preparou novidades especiais: melhorias e novos serviços desenvolvidos com foco em você.",
  "Mais melhorias chegaram na Cruz Azul! Trabalhamos constantemente para entregar sempre o melhor."
];


// ---------------- Config ----------------
const ARQUIVO_XLSX = "testes-numeros.xlsx";

// Intervalos bem mais lentos e humanos
const PAUSA_MIN_MS = 25000;   // mínimo 25s
const PAUSA_MAX_MS = 55000;   // máximo 55s

// Pausas maiores e mais aleatórias entre 80 e 160 envios
const PAUSA_SEGURANCA_A_CADA = Math.floor(Math.random() * (160 - 80 + 1)) + 80;

// Pausa de segurança entre 3min e 6min
const PAUSA_SEGURANCA_MS = Math.floor(Math.random() * (360000 - 180000 + 1)) + 180000;

// ---------------- Helpers ----------------
function delay(ms) {
  return new Promise((res) => setTimeout(res, ms));
}

function delayRandom(min, max) {
  return delay(Math.floor(Math.random() * (max - min + 1) + min));
}

function numeroValido(numero) {
  return /^[0-9]{10,13}$/.test(numero);
}

function dataHoraLocal() {
  const agora = new Date();
  return `${agora.toLocaleDateString("pt-BR")} ${agora.toLocaleTimeString("pt-BR")}`;
}

// ---------------- Carregar planilha ----------------
if (!fs.existsSync(ARQUIVO_XLSX)) {
  console.error(`Arquivo não encontrado: ${ARQUIVO_XLSX}`);
  process.exit(1);
}

const workbook = xlsx.readFile(ARQUIVO_XLSX);
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const dados = xlsx.utils.sheet_to_json(sheet);

console.log(`Planilha carregada! Total de contatos: ${dados.length}`);

// ---------------- Preparar resultados ----------------
const enviados = [];
const falhados = [];

// ---------------- ProgressBar ----------------
const progressBar = new cliProgress.SingleBar({}, cliProgress.Presets.shades_classic);

// ---------------- Inicializar client WhatsApp ----------------
const client = new Client({
  authStrategy: new LocalAuth({ clientId: "automacao-node" }),
  puppeteer: {
    headless: false,
    executablePath: "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
    args: ["--no-sandbox", "--disable-setuid-sandbox"],
  },
});

client.on("qr", () => console.log("Escaneie o QR Code..."));

client.on("ready", async () => {
  console.log("WhatsApp conectado! Iniciando envios...\n");

  console.log("Aguardando comportamento humano inicial...");
  await delayRandom(5000, 15000);

  progressBar.start(dados.length, 0);

  let contador = 0;

  for (let i = 0; i < dados.length; i++) {
    const contato = dados[i];

    const beneficiario = contato.Coluna2 || "Beneficiário";

    let celular = String(contato.Coluna3 || "").replace(/\D/g, "").trim();
    if (celular.startsWith("0")) celular = celular.replace(/^0+/, "");

    if (!numeroValido(celular)) {
      falhados.push({
        beneficiario,
        celular,
        motivo: "número inválido",
        Data_Hora: dataHoraLocal(),
      });

      progressBar.update(i + 1);
      continue;
    }

    const numeroSemPrefixo = `55${celular}`;
    const numeroZapId = `${numeroSemPrefixo}@c.us`;

    const videoSelecionado = videos[Math.floor(Math.random() * videos.length)];
    const legendaAleatoria = mensagensAleatorias[Math.floor(Math.random() * mensagensAleatorias.length)];

    const mensagemFinal = `Olá ${beneficiario}, tudo bem?\n${legendaAleatoria}`.trim();

    try {
      let registrado = true;

      if (typeof client.isRegisteredUser === "function") {
        try {
          registrado = await client.isRegisteredUser(numeroSemPrefixo);
        } catch (_) {
          await delayRandom(1000, 2000);
          registrado = await client.isRegisteredUser(numeroSemPrefixo);
        }
      }

      if (!registrado) {
        falhados.push({
          beneficiario,
          celular: numeroSemPrefixo,
          motivo: "não registrado",
          Data_Hora: dataHoraLocal(),
        });
      } else {
        // ---------------- Simular comportamento humano ----------------
        await client.sendPresenceAvailable();
        await delayRandom(800, 2000);

        await client.sendPresenceTyping(numeroZapId);
        await delayRandom(2000, 3500);

        if (Math.random() < 0.03) {
          console.log("Simulando comportamento humano extra...");
          await delayRandom(20000, 45000);
        }

        // ------------- PAUSA HUMANA EXTRA (5% de chance) -------------
        if (Math.random() < 0.05) {
          console.log("Pausa humana aleatória ativada...");
          await delayRandom(60000, 180000); // 1–3 min
        }

        // ---------------- Alternância de formato ----------------
        const modoEnvio = Math.random();

        if (modoEnvio < 0.70) {
          await client.sendMessage(numeroZapId, videoSelecionado, {
            caption: mensagemFinal.slice(0, 140),
          });

        } else if (modoEnvio < 0.90) {
          await client.sendMessage(numeroZapId, mensagemFinal);
          await delayRandom(2500, 4500);
          await client.sendMessage(numeroZapId, videoSelecionado);

        } else {
          await client.sendMessage(numeroZapId, mensagemFinal);
        }

        enviados.push({
          beneficiario,
          celular: numeroSemPrefixo,
          status: "enviado",
          Data_Hora: dataHoraLocal(),
        });
      }
    } catch (err) {
      falhados.push({
        beneficiario,
        celular: numeroSemPrefixo,
        motivo: err.message || "erro desconhecido",
        Data_Hora: dataHoraLocal(),
      });
    }

    contador++;
    progressBar.update(i + 1);

    // ---------------- Pausa de segurança ----------------
    if (contador % PAUSA_SEGURANCA_A_CADA === 0) {
      const pausa = PAUSA_SEGURANCA_MS + Math.floor(Math.random() * 60000);
      console.log(`\nPausa de segurança: aguardando ${Math.floor(pausa / 1000)} segundos...\n`);
      await delay(pausa);
    }

    // ---------------- Pausa aleatória entre envios ----------------
    await delayRandom(PAUSA_MIN_MS, PAUSA_MAX_MS);
  }

  progressBar.stop();

  // ---------------- Salvar relatórios ----------------
  const wbEnviados = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(wbEnviados, xlsx.utils.json_to_sheet(enviados), "Enviados");
  xlsx.writeFile(wbEnviados, "enviados.xlsx");

  const wbFalhados = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(wbFalhados, xlsx.utils.json_to_sheet(falhados), "Falhados");
  xlsx.writeFile(wbFalhados, "falhados.xlsx");

  console.log("\nProcesso finalizado!");
  console.log(`Total enviados: ${enviados.length}`);
  console.log(`Total falhados: ${falhados.length}`);
});

client.initialize();
