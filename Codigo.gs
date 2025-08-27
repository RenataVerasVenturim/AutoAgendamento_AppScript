/*
  Direitos Autorais (c) 2025 Renata Veras Venturim

  Licença MIT

  A permissão é concedida, gratuitamente, a qualquer pessoa que obtenha uma cópia
  deste software e dos arquivos de documentação associados (o "Software"), para negociar
  no Software sem restrições, incluindo, sem limitação, os direitos de uso, cópia,
  modificação, fusão, publicação, distribuição, sublicenciamento e/ou venda de cópias do Software,
  e de permitir que as pessoas a quem o Software é fornecido o façam, sujeitas às seguintes condições:

  O aviso de direitos autorais acima e este aviso de permissão devem ser incluídos em todas as cópias
  ou partes substanciais do Software.

  **Atribution Clause:**
  Os usuários deste software são obrigados a manter os créditos originais e as atribuições do projeto
  em quaisquer cópias ou derivações do Software.

  O SOFTWARE É FORNECIDO "COMO ESTÁ", SEM GARANTIA DE QUALQUER TIPO, EXPRESSA OU IMPLÍCITA,
  INCLUINDO, MAS NÃO SE LIMITANDO ÀS GARANTIAS DE COMERCIALIZAÇÃO, ADEQUAÇÃO A UM PROPÓSITO ESPECÍFICO
  E NÃO VIOLAÇÃO. EM NENHUM CASO OS AUTORES OU DETENTORES DOS DIREITOS AUTORAIS SERÃO RESPONSÁVEIS
  POR QUALQUER REIVINDICAÇÃO, DANOS OU OUTRAS RESPONSABILIDADES, SEJA EM AÇÃO DE CONTRATO, DELITO
  OU DE OUTRA FORMA, DECORRENTES DE, OU EM CONEXÃO COM O SOFTWARE OU O USO OU OUTRAS NEGOCIAÇÕES NO PROGRAMAS.
*/


function doGet(e) {
  const user = Session.getEffectiveUser().getEmail();
  // verificação de retorno de e mail login apenas uff -> return HtmlService.createHtmlOutput("Usuário logado: " + user);
  if (!user || !/@id\.uff\.br$/.test(user)) {
    return HtmlService.createHtmlOutput("Acesso restrito. Faça login com sua conta institucional.");
  }
  const template = HtmlService.createTemplateFromFile("Principal1");
  return template.evaluate()
    .setTitle("UFF")
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .addMetaTag("viewport", "width=device-width,initial-scale=1");
}


function Chamar(Arquivo){

  return HtmlService.createHtmlOutputFromFile(Arquivo).getContent();
}
function getCalendarId() {
  return PropertiesService.getScriptProperties().getProperty("CALENDAR_ID");
}
