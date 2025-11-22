const xlsx = require('xlsx');
const { Client, LocalAuth } = require('whatsapp-web.js');
const fs = require('fs');
const cliProgress = require('cli-progress');

// ---------------- Config ----------------
const ARQUIVO_XLSX = 'testes-numeros.xlsx'; // nome correto do arquivo
const PAUSA_MIN_MS = 1500;
const PAUSA_MAX_MS = 3000;
const PAUSA_SEGURANCA_A_CADA = 300;
const PAUSA_SEGURANCA_MS = 60000;

// ---------------- Helpers ----------------
function delay(ms) {
  return new Promise(res => setTimeout(res, ms));
}
function delayRandom(min, max) {
  return delay(Math.floor(Math.random() * (max - min + 1) + min));
}
function numeroValido(numero) {
  // 10 a 13 dígitos (DDD + número) sem + ou outros símbolos
  return /^[0-9]{10,13}$/.test(numero);
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
    args: ['--no-sandbox', '--disable-setuid-sandbox']
  }
});

client.on('qr', qr => {
  console.log('Escaneie o QR Code...');
});

client.on('ready', async () => {
  console.log('WhatsApp conectado! Iniciando envios...\n');
  progressBar.start(dados.length, 0);

  let contador = 0;

  for (let i = 0; i < dados.length; i++) {
    const contato = dados[i];
    const beneficiario = contato.BENEFICIARIO || contato.Beneficiario || 'Beneficiário';
    let celular = String(contato.CELULAR || contato.Celular || contato.telefone || '').replace(/\D/g, "");
    if (celular.startsWith('0')) celular = celular.replace(/^0+/, '');

    // valida
    if (!numeroValido(celular)) {
      console.log(`Número inválido: ${beneficiario} (${celular})`);
      falhados.push({ beneficiario, celular, motivo: 'número inválido' });
      progressBar.update(i + 1);
      continue;
    }

    const numeroSemPrefixo = `55${celular}`; // DDI +55
    const numeroZapId = `${numeroSemPrefixo}@c.us`;

    const mensagem = `
Olá ${beneficiario}, tudo bem?

Este é um aviso automático referente ao seu cadastro. Por favor, confirme o recebimento desta mensagem.

Att,
Equipe de Suporte.
    `.trim();

    try {
      // 1) Se a lib disponibiliza isRegisteredUser (verifica se o número tem WhatsApp)
      let registrado = true;
      if (typeof client.isRegisteredUser === 'function') {
        try {
          registrado = await client.isRegisteredUser(numeroSemPrefixo);
        } catch (err) {
          // se der erro, vamos tentar enviar e capturar
          registrado = true;
        }
      }

      if (!registrado) {
        console.log(`Número não registrado no WhatsApp: ${beneficiario} (${numeroSemPrefixo})`);
        falhados.push({ beneficiario, celular: numeroSemPrefixo, motivo: 'não registrado' });
      } else {
        // 2) envia com a API da lib (mais robusto)
        const sent = await client.sendMessage(numeroZapId, mensagem);

        // se chegou aqui sem exception, consideramos enviado
        console.log(`Enviado → ${beneficiario} (${numeroSemPrefixo})`);
        enviados.push({ beneficiario, celular: numeroSemPrefixo, status: 'enviado', timestamp: new Date().toISOString() });
      }
    } catch (err) {
      // captura mensagens não entregues ou erros da lib
      console.log(`Falha ao enviar → ${beneficiario} (${numeroSemPrefixo}) - ${err.message || err}`);
      falhados.push({ beneficiario, celular: numeroSemPrefixo, motivo: err.message || 'erro desconhecido' });
    }

    contador++;
    progressBar.update(i + 1);

    // Pausa de segurança a cada N envios
    if (contador % PAUSA_SEGURANCA_A_CADA === 0) {
      console.log(`\nPausa de segurança: aguardando ${PAUSA_SEGURANCA_MS/1000} segundos...\n`);
      await delay(PAUSA_SEGURANCA_MS);
    }

    // pausa aleatória entre envios
    await delayRandom(PAUSA_MIN_MS, PAUSA_MAX_MS);
  }

  progressBar.stop();

  // salvar relatórios
  const wbEnviados = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(wbEnviados, xlsx.utils.json_to_sheet(enviados), 'Enviados');
  xlsx.writeFile(wbEnviados, 'enviados.xlsx');

  const wbFalhados = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(wbFalhados, xlsx.utils.json_to_sheet(falhados), 'Falhados');
  xlsx.writeFile(wbFalhados, 'falhados.xlsx');

  console.log('\nProcesso finalizado!');
  console.log(`Total enviados: ${enviados.length}`);
  console.log(`Total falhados: ${falhados.length}`);

  // não fecha o browser automaticamente para você revisar, descomente se quiser fechar
  // await client.destroy();
});

client.initialize();
