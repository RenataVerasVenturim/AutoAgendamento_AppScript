/* 
Direitos Autorais (c) 2025 Renata Veras Venturim
Licença MIT

A permissão é concedida, gratuitamente, a qualquer pessoa que obtenha uma cópia deste software
e dos arquivos de documentação associados (o "Software"), para negociar no Software sem restrições,
incluindo, sem limitação, os direitos de uso, cópia, modificação, fusão, publicação, distribuição,
sublicenciamento e/ou venda de cópias do Software, e permitir que outros o façam, sujeitas às seguintes condições:

O aviso de direitos autorais acima e este aviso de permissão devem ser incluídos em todas as cópias
ou partes substanciais do Software.

**Atribution Clause:** Os usuários deste software são obrigados a manter os créditos originais e as atribuições
do projeto em quaisquer cópias ou derivações do Software.

O SOFTWARE É FORNECIDO "COMO ESTÁ", SEM GARANTIA DE QUALQUER TIPO, EXPRESSA OU IMPLÍCITA,
INCLUINDO, MAS NÃO SE LIMITANDO ÀS GARANTIAS DE COMERCIALIZAÇÃO, ADEQUAÇÃO A UM PROPÓSITO ESPECÍFICO E NÃO VIOLAÇÃO.
EM NENHUM CASO OS AUTORES OU DETENTORES DOS DIREITOS AUTORAIS SERÃO RESPONSÁVEIS POR QUALQUER REIVINDICAÇÃO,
DANOS OU OUTRAS RESPONSABILIDADES, SEJA EM AÇÃO DE CONTRATO, DELITO OU DE OUTRA FORMA, DECORRENTES DE,
OU EM CONEXÃO COM O SOFTWARE OU O USO OU OUTRAS NEGOCIAÇÕES NO PROGRAMA.
*/

function doGet(e) {
  const user = Session.getActiveUser().getEmail();
  if (!user || !/@id\.uff\.br$/i.test(user)) {
    return HtmlService.createHtmlOutput("Acesso restrito. Faça login com sua conta institucional id.uff.br.");
  }

  const action = e && e.parameter && e.parameter.action;
  const token  = e && e.parameter && e.parameter.token;

  // 🔹 Rota para abrir Termos de Privacidade
  if (e && e.parameter && e.parameter.page === "TermosPrivacidade") {
    return HtmlService.createTemplateFromFile("TermosPrivacidade")
      .evaluate()
      .setTitle("Termos de Privacidade")
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  }

  // 🔹 Rota de ações (ex.: cancelar agendamento)
  if (action && token) {
    try {
      _assertOwnerByToken_(token);

      const t = HtmlService.createTemplateFromFile("Acoes");
      t.userEmail = user.toUpperCase();
      t.token = token;
      t.action = String(action).toLowerCase();
      return t.evaluate()
        .setTitle("UFF - Ações do Agendamento")
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .addMetaTag("viewport", "width=device-width,initial-scale=1");
    } catch (err) {
      return HtmlService.createHtmlOutput(`<h3>${err.message}</h3>`);
    }
  }

  // 🔹 Página principal
  const template = HtmlService.createTemplateFromFile("Principal1");
  template.userEmail = user.toUpperCase();
  return template.evaluate()
    .setTitle("UFF")
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .addMetaTag("viewport", "width=device-width,initial-scale=1");
}

// Função para incluir arquivos HTML
function Chamar(Arquivo) {
  return HtmlService.createHtmlOutputFromFile(Arquivo).getContent();
}

// Função para recuperar o ID do calendário
function getCalendarId() {
  return PropertiesService.getScriptProperties().getProperty("CALENDAR_ID");
}

/* ===========================
   Exemplo de função comentada para dados do usuário
=========================== */
// function VerDadosUsuario() {
//   return {
//     UserActive: Session.getActiveUser().getEmail(),
//     UserEffective: Session.getEffectiveUser().getEmail(),
//   };
// }
