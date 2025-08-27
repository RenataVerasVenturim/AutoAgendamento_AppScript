/*
  Direitos Autorais (c) 2025 Renata Veras Venturim
  Licença MIT
*/

const calendarId = "primary";
const SHEET_CONFIG = "Config_Agendamento";

/**
 * Lê a configuração de horários do Sheets
 */
function getConfiguracoes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_CONFIG);
  const data = sheet.getDataRange().getValues();
  const config = [];

  for (let i = 1; i < data.length; i++) {
    config.push({
      dia: data[i][0],                       // Segunda, Terça, etc.
      inicio: data[i][1],                    // HH:mm
      fim: data[i][2],                       // HH:mm
      intervalo: parseInt(data[i][3]) || 60, // intervalo em minutos
      ativo: data[i][4] === "SIM"            // true / false
    });
  }
  return config;
}

/**
 * Retorna os dias da semana permitidos
 */
function getDiasPermitidos() {
  const config = getConfiguracoes();
  const dias = [];
  const mapa = { "Domingo":0, "Segunda":1, "Terça":2, "Quarta":3, "Quinta":4, "Sexta":5, "Sábado":6 };
  config.forEach(c => {
    if(c.ativo && mapa[c.dia] !== undefined){
      dias.push(mapa[c.dia]);
    }
  });
  return dias;
  
}

/**
 * Converte string YYYY-MM-DD para objeto Date
 */
function parseDataYYYYMMDD(dataStr) {
  if (!dataStr) return null;
  const partes = dataStr.split("-");
  if (partes.length !== 3) return null;
  const ano = parseInt(partes[0], 10);
  const mes = parseInt(partes[1], 10) - 1; // JS Date: 0 = Janeiro
  const dia = parseInt(partes[2], 10);
  return { dia, mes, ano };
}

/**
 * Calcula os horários disponíveis
 */

/**
 * Converte valor de hora do Sheets para string "HH:mm"
 * Aceita string "09:00" ou número 0.375 (Excel/Sheets)
 */
function formatHoraString(valor) {
  if (typeof valor === 'string') return valor;
  if (typeof valor === 'number') { // hora como número do Sheets
    const date = new Date(1899, 11, 30); // base Excel
    date.setHours(Math.floor(valor * 24));
    date.setMinutes(Math.round((valor * 24 * 60) % 60));
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "HH:mm");
  }
  return "00:00";
}

/**
 * Retorna os horários disponíveis para uma data
 */
function getHorariosDisponiveis(dataSelecionada) {
  const config = getConfiguracoes();
  const dataParts = parseDataYYYYMMDD(dataSelecionada);
  if (!dataParts) return { horarios: [], debug: "Data inválida" };

  const { dia, mes, ano } = dataParts;
  const data = new Date(ano, mes, dia);
  const diaSemanaStr = ["Domingo","Segunda","Terça","Quarta","Quinta","Sexta","Sábado"][data.getDay()];

  // Busca regra da planilha
  const regra = config.find(c => c.dia === diaSemanaStr && c.ativo);
  if (!regra) return { horarios: [], debug: "Dia não disponível" };
  if (!regra.inicio || !regra.fim) return { horarios: [], debug: "Início ou fim não configurado" };

  const inicioStr = formatHoraString(regra.inicio);
  const fimStr = formatHoraString(regra.fim);

  const [horaInicio, minInicio] = inicioStr.split(":").map(Number);
  const [horaFim, minFim] = fimStr.split(":").map(Number);

  // Gerar todos os horários possíveis
  let horaAtual = new Date(ano, mes, dia, horaInicio, minInicio);
  const horaFinal = new Date(ano, mes, dia, horaFim, minFim);

  const horariosPermitidos = [];
  while (horaAtual < horaFinal) {
    horariosPermitidos.push(Utilities.formatDate(horaAtual, Session.getScriptTimeZone(), "HH:mm"));
    horaAtual = new Date(horaAtual.getTime() + regra.intervalo * 60000);
  }

  // Busca eventos no Calendar
  
  const inicioDia = new Date(ano, mes , dia, 0, 0, 0);
  const fimDia = new Date(ano, mes , dia, 23, 59, 59);

  const eventos = Calendar.Events.list(calendarId, {
    timeMin: inicioDia.toISOString(),
    timeMax: fimDia.toISOString(),
    singleEvents: true,
    orderBy: "startTime"
  }).items || [];

  // Filtra eventos de hora marcada (ignora all-day)
  const horariosOcupados = eventos.map(e => {
    if (e.start.date && e.end.date) return null; // ignora all-day
    const inicio = e.start.dateTime ? new Date(e.start.dateTime) : new Date(e.start.date);
    const fim = e.end.dateTime ? new Date(e.end.dateTime) : new Date(e.end.date);
    return { inicio, fim };
  }).filter(Boolean);

  // Filtra horários livres
  const horariosLivres = horariosPermitidos.filter(h => {
    const [hora, min] = h.split(":").map(Number);
    const slotInicio = new Date(ano, mes , dia, hora, min);
    const slotFim = new Date(slotInicio.getTime() + regra.intervalo * 60000);

    return !horariosOcupados.some(ev => slotInicio < ev.fim && slotFim > ev.inicio);
  });

  return {
    horarios: horariosLivres,
    debug: {
      diaSemana: diaSemanaStr,
      regra,
      horariosPermitidos,
      eventosDoDia: eventos.length
    }
  };
}
//rating limiting (limite por tempo por usuário)--------------------------//

  //registrar tentativas maliciosas
function registrarTentativa(email, mensagem, tipo = "INFO") {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Logs_Seguranca");
  if (!sheet) return;
  
  const userKey = Session.getTemporaryActiveUserKey() || "N/A"; // identifica usuário anonimamente
  
  sheet.appendRow([
    new Date(),      // Data e hora
    userKey,         // ID temporário do usuário
    email || "Desconhecido", 
    tipo,            // Tipo de evento: INFO, ERRO, RATE_LIMIT, SUCESSO
    mensagem         // Mensagem descritiva
  ]);
}

function canUserProceedWithLog(email) {
  const userKey = Session.getTemporaryActiveUserKey();
  const userProps = PropertiesService.getUserProperties();
  const maxRequests = 5;      // Ajuste conforme necessidade
  const windowMinutes = 30;

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
//verificar usuário autenticação
function verificarUsuarioUFF() {
  const user = Session.getEffectiveUser().getEmail();
  if (!user || !/@id\.uff\.br$/.test(user)) {
    throw new Error("Acesso restrito. Faça login com sua conta institucional.");
  }
  return user;
}


//---------------------------------------------------------------------------//
function processarAgendamento(dados) {
  try {
    const user = verificarUsuarioUFF(); // já lança erro se não autorizado
    if (!canUserProceedWithLog(dados.email)) {
      return { sucesso: false, mensagem: "Você excedeu o limite de requisições. Tente novamente mais tarde." };
    }

  const regexUFF = /^[a-zA-Z0-9._%+-]+@ID\.UFF\.BR$/;
  if (!regexUFF.test(dados.email) || dados.email.length>100) {
    return { sucesso: false, mensagem: "Apenas e-mails válidos do domínio @id.uff.br são permitidos." 
    };
  }

  // Validação do nome
  if (!dados.nome || dados.nome.length < 3 || dados.nome.length>50) {
    return { sucesso: false, mensagem: "Nome inválido." 
    };
  }

  // Validação da data
  const regexData = /^\d{4}-\d{2}-\d{2}$/;
  if (!regexData.test(dados.data)) {
    return{sucesso:false,mensagem:"Data inválida. Selecione uma data válida (formato: YYYY-MM-DD)."
    };
  };

  const horariosDisponiveis = getHorariosDisponiveis(dados.data).horarios;

  if (!horariosDisponiveis.includes(dados.hora)) {
    return { sucesso: false, mensagem: "Horário já reservado ou indisponível. Por favor, selecione outro." };
  }

  // impedir xss
  function escaparHTML(str) {
  if (!str) return "";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

  // Cria os objetos de data/hora para início e fim
  const inicio = new Date(`${dados.data}T${dados.hora}:00`);
  const fim = new Date(inicio.getTime() + Number(dados.duracao) * 60000);

  // Cria o evento com link do Google Meet
  const evento = Calendar.Events.insert(
    {
      summary: dados.nome + '-' +dados.unidade,
      description: escaparHTML(dados.obs),
      start: { dateTime: inicio.toISOString(), timeZone: Session.getScriptTimeZone() },
      end: { dateTime: fim.toISOString(), timeZone: Session.getScriptTimeZone() },
      attendees: [{ email: dados.email }],
      conferenceData: {
        createRequest: { requestId: Utilities.getUuid() }
      }
    },
    calendarId,
    { conferenceDataVersion: 1 }
  );

  const meetLink = evento.conferenceData?.entryPoints?.[0]?.uri || "Link do Meet não disponível.";

  // ✅ Grava na planilha "Agendamentos"
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Agendamentos");
  if (sheet) {
    sheet.appendRow([
      escaparHTML(dados.nome),
      escaparHTML(dados.unidade),
      escaparHTML(dados.email),
      escaparHTML(dados.data),
      escaparHTML(dados.hora),
      escaparHTML(dados.duracao),
      escaparHTML(dados.obs),
      escaparHTML(meetLink) // opcional: gravar link do Meet
    ]);
  }
// === Enviar e-mail de confirmação ===
try {
  const assunto = "Confirmação de Agendamento - UFF";
  const corpo = `
    Olá, ${escaparHTML(dados.nome)}!
    
    Seu agendamento foi realizado com sucesso.

    📌 **Detalhes do Agendamento:**
    • Unidade: ${escaparHTML(dados.unidade)}
    • Data: ${escaparHTML(dados.data)}
    • Horário: ${escaparHTML(dados.hora)}
    • Duração: ${escaparHTML(dados.duracao)} minutos
    • Assunto: ${escaparHTML(dados.obs)}

    🔗 Link da reunião (Google Meet): ${escaparHTML(meetLink)}

    Caso não tenha solicitado este agendamento, entre em contato com a nossa equipe.

    ---
    Universidade Federal Fluminense - PROPPI
    Este é um e-mail automático, por favor, não responda.
`;

  MailApp.sendEmail({
    to: dados.email,
    subject: assunto,
    htmlBody: corpo
  });

} catch (erroEmail) {
  Logger.log("Falha ao enviar e-mail: " + erroEmail.message);
}

  return {
    sucesso: true,
    mensagem: `Seu agendamento foi registrado com sucesso!<br>
               <b>Data:</b> ${dados.data} às ${dados.hora}<br>
               <b>Link da reunião:</b> <a href="${meetLink}" target="_blank">${meetLink}</a>`
  };
}   catch (err) {
    return { sucesso: false, mensagem: err.message };
  }
}



