const XLSX = require("xlsx");
const puppeteer = require("puppeteer");
const fs = require("fs");
const cliProgress = require("cli-progress");

// Função aleatória para pausas
function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function delayRandom(min, max) {
  return delay(Math.floor(Math.random() * (max - min + 1) + min));
}

// Validação básica de número
function numeroValido(numero) {
  const regex = /^[0-9]{10,13}$/;
  return regex.test(numero);
}

async function enviarMensagens() {

  // Carregar XLSX
  const workbook = XLSX.readFile("testes-numeros.xlsx");
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const contatos = XLSX.utils.sheet_to_json(sheet);

  console.log(`Planilha carregada! Total de contatos: ${contatos.length}\n`);

  // Listas de resultados
  const enviados = [];
  const falhados = [];

  // Barra de progresso
  const bar = new cliProgress.SingleBar({}, cliProgress.Presets.shades_classic);
  bar.start(contatos.length, 0);

  // Iniciar WhatsApp
  const browser = await puppeteer.launch({
    headless: false,
    args: ["--no-sandbox"]
  });

  const page = await browser.newPage();
  await page.goto("https://web.whatsapp.com");

  console.log("Escaneie o QR Code...");

  await page.waitForSelector("canvas", { timeout: 60000 }).catch(() => {});

  console.log("Aguardando login...");

  // 🔥 NOVO SISTEMA DE DETECÇÃO DE LOGIN — 100% COMPATÍVEL
  await page.waitForFunction(() => {
    return (
      document.querySelector('[title="Nova conversa"]') ||
      document.querySelector('[aria-label="Nova conversa"]') ||
      document.querySelector('[aria-label="Nova mensagem"]') ||
      document.querySelector('[aria-label="Iniciar nova conversa"]') ||
      document.querySelector("div[role='textbox']")
    );
  }, { timeout: 0 });

  console.log("Login realizado. Iniciando envios...\n");

  let contadorEnvios = 0;

  for (const contato of contatos) {
    const beneficiario = contato.BENEFICIARIO;
    const celular = String(contato.CELULAR).replace(/\D/g, "");

    // Validação do número
    if (!numeroValido(celular)) {
      console.log(`Número inválido → ${beneficiario} (${celular})`);
      falhados.push({ beneficiario, celular, motivo: "número inválido" });
      bar.increment();
      continue;
    }

    // Mensagem
    const msg = `
Olá ${beneficiario}, tudo bem?

Esse é meu primeiro teste de automação; espero que dê certo e gostaria de receber seu feedback.

Estou entrando em contato para informar que, se recebeu essa mensagem, foi contemplado com um grande prêmio...
A Telemática adverte: não caia em golpes! Nunca forneça seus dados pessoais ou bancários a terceiros.
KKKK

Atenciosamente,
Marcelo Martinho

Qualquer dúvida, estamos à disposição!
    `.trim();

    try {
      const url = `https://web.whatsapp.com/send?phone=${celular}&text=${encodeURIComponent(msg)}`;
      await page.goto(url);

      await delayRandom(3000, 6000);

      await page.waitForSelector('span[data-icon="send"]', { timeout: 8000 });
      await page.click('span[data-icon="send"]');

      console.log(`Mensagem enviada → ${beneficiario} (${celular})`);
      enviados.push({ beneficiario, celular, status: "enviado" });

    } catch (e) {
      console.log(`Falha ao enviar → ${beneficiario} (${celular})`);
      falhados.push({ beneficiario, celular, motivo: "erro ao enviar" });
    }

    contadorEnvios++;
    bar.increment();

    if (contadorEnvios % 300 === 0) {
      console.log("\nPausa de segurança de 1 minuto…\n");
      await delay(60000);
    }

    await delayRandom(2000, 4000);
  }

  bar.stop();

  console.log("\nGerando relatórios...");

  const wbEnviados = XLSX.utils.book_new();
  const wsEnviados = XLSX.utils.json_to_sheet(enviados);
  XLSX.utils.book_append_sheet(wbEnviados, wsEnviados, "Enviados");
  XLSX.writeFile(wbEnviados, "enviados.xlsx");

  const wbFalhados = XLSX.utils.book_new();
  const wsFalhados = XLSX.utils.json_to_sheet(falhados);
  XLSX.utils.book_append_sheet(wbFalhados, wsFalhados, "Falhados");
  XLSX.writeFile(wbFalhados, "falhados.xlsx");

  console.log("\nProcesso finalizado!");
  console.log(`Total enviados: ${enviados.length}`);
  console.log(`Total falhados: ${falhados.length}`);

  // await browser.close();
}

enviarMensagens();
