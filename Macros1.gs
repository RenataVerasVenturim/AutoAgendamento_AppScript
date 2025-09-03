/*
Licença MIT

Copyright (c) 2025 RENATA VERAS VENTURIM

MIT License

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

**Atribution Clause:**
Users of this software are required to maintain the original credits and attributions of the project in any copies or derivatives of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
*/


const calendarId = "primary";
const SHEET_CONFIG = "Config_Agendamento";

/* ===========================
   Verifica se o usuário pode agendar
   com base na aba "Controle_Certificados"
   Coluna B = e-mail, Coluna C = status
=========================== */
function podeAgendar(email) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Controle_Certificados");
  if (!sheet) throw new Error("A aba 'Controle_Certificados' não foi encontrada.");

  const dados = sheet.getDataRange().getValues();

  for (let i = 1; i < dados.length; i++) {
    const emailPlanilha = String(dados[i][1] || "").trim().toLowerCase();
    const status = String(dados[i][2] || "").trim().toUpperCase();
    if (emailPlanilha === email.toLowerCase() && status === "PRONTO") return true;
  }

  return false;
}

/* ===========================
   Lê a configuração de horários do Sheets
=========================== */
/**
 * Busca configurações de uma aba específica do Sheets.
 * Se a aba for "Config_Agendamento", retorna array de objetos {dia, inicio, fim, intervalo, ativo}
 * Se a aba for "Config_Feriados", retorna array de strings YYYY-MM-DD
 */
function getConfigPlanilha(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();

  if (sheetName === SHEET_CONFIG) {
    // Configuração de horários
    const config = [];
    for (let i = 1; i < data.length; i++) {
      config.push({
        dia: data[i][0],                 // Segunda, Terça, etc.
        inicio: data[i][1],              // HH:mm
        fim: data[i][2],                 // HH:mm
        intervalo: parseInt(data[i][3]) || 60, // minutos
        ativo: data[i][4] === "SIM"      // true/false
      });
    }
    return config;
  } else if (sheetName === "Config_Feriados") {
    // Feriados
    return data
      .slice(1) // Ignora cabeçalho
      .flat()
      .filter(d => d)
      .map(d => (d instanceof Date) ? Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd") : d);
  }

  return data; // Retorna cru se aba diferente
}


/* ===========================
   Retorna os dias da semana permitidos
=========================== */
function getDiasPermitidos() {
  const config = getConfigPlanilha(SHEET_CONFIG);
  const dias = [];
  const mapa = {
    "Domingo": 0, "Segunda": 1, "Terça": 2, "Quarta": 3,
    "Quinta": 4, "Sexta": 5, "Sábado": 6
  };

  config.forEach(c => {
    if (c.ativo && mapa[c.dia] !== undefined) dias.push(mapa[c.dia]);
  });

  return dias;
}

/* ===========================
   Converte string YYYY-MM-DD para objeto Date
=========================== */
function parseDataYYYYMMDD(dataStr) {
  if (!dataStr) return null;
  const partes = dataStr.split("-");
  if (partes.length !== 3) return null;

  const ano = parseInt(partes[0], 10);
  const mes = parseInt(partes[1], 10) - 1; // JS Date: 0 = Janeiro
  const dia = parseInt(partes[2], 10);

  return { dia, mes, ano };
}

/* ===========================
   Converte valor de hora do Sheets para string "HH:mm"
=========================== */
function formatHoraString(valor) {
  if (typeof valor === 'string') return valor;
  if (typeof valor === 'number') {
    const date = new Date(1899, 11, 30); // base Excel date
    date.setHours(Math.floor(valor * 24));
    date.setMinutes(Math.round((valor * 24 * 60) % 60));
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "HH:mm");
  }
  return "00:00";
}

/* ===========================
   Retorna os horários disponíveis para uma data
=========================== */
function getHorariosDisponiveis(dataSelecionada) {
  const config = getConfigPlanilha(SHEET_CONFIG);
  const dataParts = parseDataYYYYMMDD(dataSelecionada);
  if (!dataParts) return { horarios: [], debug: "Data inválida" };

  const { dia, mes, ano } = dataParts;
  const data = new Date(ano, mes, dia);
  const diaSemanaStr = ["Domingo","Segunda","Terça","Quarta","Quinta","Sexta","Sábado"][data.getDay()];
  const regra = config.find(c => c.dia === diaSemanaStr && c.ativo);

  if (!regra) return { horarios: [], debug: "Dia não disponível" };

  const inicioStr = formatHoraString(regra.inicio);
  const fimStr = formatHoraString(regra.fim);

  const [horaInicio, minInicio] = inicioStr.split(":").map(Number);
  const [horaFim, minFim] = fimStr.split(":").map(Number);

  let horaAtual = new Date(ano, mes, dia, horaInicio, minInicio);
  const horaFinal = new Date(ano, mes, dia, horaFim, minFim);

  const horariosPermitidos = [];
  while (horaAtual < horaFinal) {
    horariosPermitidos.push(Utilities.formatDate(horaAtual, Session.getScriptTimeZone(), "HH:mm"));
    horaAtual = new Date(horaAtual.getTime() + regra.intervalo * 60000);
  }

  // Busca eventos já agendados no dia
  const inicioDia = new Date(ano, mes, dia, 0, 0, 0);
  const fimDia = new Date(ano, mes, dia, 23, 59, 59);
  const eventos = Calendar.Events.list(calendarId, {
    timeMin: inicioDia.toISOString(),
    timeMax: fimDia.toISOString(),
    singleEvents: true,
    orderBy: "startTime"
  }).items || [];

  const horariosOcupados = eventos.map(e => {
    if (e.start.date && e.end.date) return null;
    const inicio = e.start.dateTime ? new Date(e.start.dateTime) : new Date(e.start.date);
    const fim = e.end.dateTime ? new Date(e.end.dateTime) : new Date(e.end.date);
    return { inicio, fim };
  }).filter(Boolean);

  const horariosLivres = horariosPermitidos.filter(h => {
    const [hora, min] = h.split(":").map(Number);
    const slotInicio = new Date(ano, mes, dia, hora, min);
    const slotFim = new Date(slotInicio.getTime() + regra.intervalo * 60000);
    return !horariosOcupados.some(ev => slotInicio < ev.fim && slotFim > ev.inicio);
  });

  return {
    horarios: horariosLivres,
    duracao: regra.intervalo,
    debug: { diaSemana: diaSemanaStr, regra, horariosPermitidos, eventosDoDia: eventos.length }
  };
}

/* ===========================
   Limite de requisições por usuário
=========================== */
function registrarTentativa(email, mensagem, tipo = "INFO") {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Logs_Seguranca");
  if (!sheet) return;

  const userKey = Session.getTemporaryActiveUserKey() || "N/A";
  sheet.appendRow([
    new Date(),
    userKey,
    email || "Desconhecido",
    tipo,
    mensagem
  ]);
}

function canUserProceedWithLog(email) {
  const userKey = Session.getTemporaryActiveUserKey();
  const userProps = PropertiesService.getUserProperties();
  const maxRequests = 3;
  const windowMinutes = 1;
  const now = Date.now();

  let data = userProps.getProperty("rateLimit_" + userKey);
  if (data) {
    data = JSON.parse(data).filter(ts => now - ts < windowMinutes * 60 * 1000);
  } else {
    data = [];
  }

  if (data.length >= maxRequests) {
    registrarTentativa(email, "Rate limit atingido", "RATE_LIMIT");
    return false;
  }

  data.push(now);
  userProps.setProperty("rateLimit_" + userKey, JSON.stringify(data));
  return true;
}

/* ===========================
   Verificação de usuário institucional
=========================== */
function verificarUsuarioUFF() {
  const user = Session.getActiveUser().getEmail();
  if (!user || !/@id\.uff\.br$/.test(user)) {
    throw new Error("Acesso restrito. Faça login com sua conta institucional.");
  }
  return user;
}

// Escapar HTML para XSS
function escaparHTML(str) {
      if (!str) return "";
      return String(str)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

/* ===========================
   Processa o agendamento
=========================== */
/* ===========================
   Processa o agendamento
=========================== */
function processarAgendamento(dados) {
  try {
    const user = verificarUsuarioUFF();

    if (!podeAgendar(dados.email)) {
      return { sucesso: false, mensagem: "Seu certificado ainda não está disponível para retirada." };
    }

    if (!canUserProceedWithLog(dados.email)) {
      return { sucesso: false, mensagem: "Você excedeu o limite de requisições. Tente novamente mais tarde." };
    }

    const regexUFF = /^[a-zA-Z0-9._%+-]+@ID\.UFF\.BR$/;
    if (!regexUFF.test(dados.email) || dados.email.length > 100) {
      return { sucesso: false, mensagem: "Apenas e-mails válidos do domínio @id.uff.br são permitidos." };
    }

    if (!dados.nome || dados.nome.length < 3 || dados.nome.length > 50) {
      return { sucesso: false, mensagem: "Nome inválido." };
    }
    const regexProcesso = /^(23069\.\d{6}\/\d{4}-\d{2})?$/;
    if (!regexProcesso.test(dados.processo)) {
      return { sucesso: false, mensagem: "Número de processo inválido. Deve estar no formato 23069.xxxxxx/AAAA-xx" };
    }

    const regexData = /^\d{4}-\d{2}-\d{2}$/;
    if (!regexData.test(dados.data)) {
      return { sucesso: false, mensagem: "Data inválida. Selecione uma data válida (formato: YYYY-MM-DD)." };
    }

    const horariosDisponiveis = getHorariosDisponiveis(dados.data).horarios;
    if (!horariosDisponiveis.includes(dados.hora)) {
      return { sucesso: false, mensagem: "Horário já reservado ou indisponível. Por favor, selecione outro." };
    }

    const inicio = new Date(`${dados.data}T${dados.hora}:00`);
    const fim = new Date(inicio.getTime() + Number(dados.duracao) * 60000);

    const evento = Calendar.Events.insert({
      summary: `${dados.nome} - ${dados.unidade}`,
      description: escaparHTML(dados.obs),
      start: { dateTime: inicio.toISOString(), timeZone: Session.getScriptTimeZone() },
      end: { dateTime: fim.toISOString(), timeZone: Session.getScriptTimeZone() },
      attendees: [{ email: dados.email }],
      conferenceData: { createRequest: { requestId: Utilities.getUuid() } },
      reminders: { useDefault: false, overrides: [{ method: "email", minutes: 1440 }] }
    }, calendarId, { conferenceDataVersion: 1, sendUpdates: "all" });

    const meetLink = evento.conferenceData?.entryPoints?.[0]?.uri || "Link do Meet não disponível.";
    const eventId = evento.id;
    const agendamentoId = Utilities.getUuid();
    const token = Utilities.getUuid();
    const baseUrl = ScriptApp.getService().getUrl();
    const linkCancelar = `${baseUrl}?action=cancelar&token=${token}`;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Agendamentos");
    const consentStr = (dados.consentimento === true || String(dados.consentimento) === "true") ? "SIM" : "NÃO";
    const dataConsentimento = consentStr === "SIM" ? new Date() : ""; // só grava timestamp se consentiu

    sheet.appendRow([
      escaparHTML(dados.nome),       // Nome
      escaparHTML(dados.unidade),    // Unidade
      escaparHTML(dados.processo),   // Processo
      escaparHTML(dados.email),      // Email
      escaparHTML(dados.data),       // Data
      escaparHTML(dados.hora),       // Hora
      escaparHTML(dados.duracao),    // Tempo
      escaparHTML(dados.obs),        // Assunto           
      escaparHTML(meetLink),         // Link meet
      agendamentoId,                 // AgendamentoID (UUID)
      token,                         // Token (UUID)
      eventId,                       // EventId (id do evento no Calendar)
      "Ativo",                       // Status (“Ativo”/“Cancelado”/“Reagendado”)
      consentStr,                     // Consentimento SIM/NÃO
      Utilities.formatDate(dataConsentimento, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss") // CreatedAt
    ]);

    // Envia e-mail de confirmação
    const assunto = "Confirmação de Agendamento - UFF";
    const corpo = `
      Olá, ${escaparHTML(dados.nome)}!<br><br>
      Seu agendamento foi realizado com sucesso.<br><br>
      <b>Unidade:</b> ${escaparHTML(dados.unidade)}<br>
      <b>Processo:</b> ${escaparHTML(dados.processo)}<br>
      <b>Data:</b> ${escaparHTML(dados.data)}<br>
      <b>Horário:</b> ${escaparHTML(dados.hora)}<br>
      <b>Assunto:</b> ${escaparHTML(dados.obs)}<br>
      <b>Endereço:</b> Rua Miguel de frias, 09, 3º andar, Sala da Secretaria , Icaraí, Niterói - RJ <br>
      <b>Link da reunião:</b> <a href="${meetLink}" target="_blank">${meetLink}</a><br><br>
      Caso precise cancelar, clique <a href="${linkCancelar}" target="_blank">aqui</a>.
    `;
    MailApp.sendEmail({ to: dados.email, subject: assunto, htmlBody: corpo });

    
    return {
      sucesso: true,
      mensagem: `Seu agendamento foi registrado com sucesso!<br>
<b>Data:</b> ${dados.data} às ${dados.hora}<br>
<b>Link da reunião:</b> <a href="${meetLink}" target="_blank">${meetLink}</a>`
    };

  } catch (err) {
    Logger.log("Erro ao processar agendamento: " + err.message);
    return { sucesso: false, mensagem: err.message };
  }
}


// =========================
// Utilitários Agendamentos
// =========================

function _getAgSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Agendamentos");
  if (!sheet) throw new Error("A aba 'Agendamentos' não foi encontrada.");
  return sheet;
}
function _findRowByToken_(token) {
  const sheet = _getAgSheet_();
  const data = sheet.getDataRange().getValues();
  const IDX_TOKEN = 10; // Token (UUID)
  
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][IDX_TOKEN]) === String(token)) {
      return { rowNumber: r + 1, row: data[r] }; // +1 pois a planilha tem cabeçalho
    }
  }
  return null;
}


function _assertOwnerByToken_(token) {
  const user = Session.getActiveUser().getEmail();
  if (!user || !/@id\.uff\.br$/i.test(user)) {
    throw new Error("Acesso restrito. Faça login com sua conta institucional.");
  }
  const match = _findRowByToken_(token);
  if (!match) throw new Error("Agendamento não encontrado ou token inválido.");

  const emailAgendamento = String(match.row[3]).toLowerCase();
  if (emailAgendamento !== String(user).toLowerCase()) {
    throw new Error("Você não tem permissão para operar este agendamento.");
  }
  return { user, sheet: _getAgSheet_(), rowNumber: match.rowNumber, row: match.row };
}
function cancelarAgendamentoPorToken(token) {
  const sheet = _getAgSheet_();
  const data = sheet.getDataRange().getValues();
  const IDX_TOKEN = 10;
  const IDX_STATUS = 12;

  let linha = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][IDX_TOKEN]) === String(token)) {
      linha = i + 1;
      break;
    }
  }

  if (linha === -1) throw new Error("Agendamento não encontrado.");

  const eventId = data[linha - 1][11]; // EventId
  if (eventId) {
    try {
      Calendar.Events.remove(calendarId, eventId);
    } catch (e) {
      Logger.log("Erro ao remover evento: " + e.message);
    }
  }

  sheet.getRange(linha, IDX_STATUS + 1).setValue("Cancelado");
  // Dados para e-mail
  // Recupera os dados da linha do agendamento
  const row = sheet.getRange(linha, 1, 1, sheet.getLastColumn()).getValues()[0];
  const nomeUsuario = row[0];         // Nome
  const emailUsuario = Session.getActiveUser().getEmail(); // E-mail do usuário logado
  const baseUrl = ScriptApp.getService().getUrl();
  // Envia e-mail de confirmação de cancelamento
  MailApp.sendEmail({
    to: emailUsuario,
    subject: "Confirmação de Cancelamento de Agendamento",
  htmlBody: `
    <p>Olá, ${escaparHTML(nomeUsuario)},</p>
    <p>Seu agendamento foi <b>cancelado com sucesso</b>.</p>
    <p>Se precisar agendar novamente, acesse nossa plataforma clicando aqui: 
       <a href="${baseUrl}?page=index1.html" target="_blank">Agendar</a>
    </p>
    <p>Atenciosamente,<br>Universidade Federal Fluminense</p>
  `});
  return { sucesso: true, mensagem: "Agendamento cancelado com sucesso." };
}



