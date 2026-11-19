
const ID_PLANILHA = "1X0jjZKfoP7WqIWlVQ4mKjq8aFBkOf-Ex-0IzNdMKfF0";

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile("Index")
    .setTitle("Gerador de OS - Eletrônica");
}

function salvarOS(dados) {
  const planilha = SpreadsheetApp.openById(ID_PLANILHA);

  const clientes = planilha.getSheetByName("Clientes");
  const ordens = planilha.getSheetByName("Ordens_Servico");

  if (!clientes || !ordens) {
    throw new Error("Não foi possível encontrar as abas da planilha.");
  }

  const agora = new Date();

  const idCliente = "CLI" + Utilities.getUuid().slice(0, 8).toUpperCase();

  const numeroOS = "OS" +
    Utilities.formatDate(
      agora,
      Session.getScriptTimeZone(),
      "yyyyMMddHHmmss"
    );

  clientes.appendRow([
    idCliente,
    dados.nome,
    dados.telefone,
    agora
  ]);

  ordens.appendRow([
    numeroOS,
    idCliente,
    agora,
    dados.marca,
    dados.polegadas,
    dados.diagnostico,
    dados.servico,
    Number(dados.valor) || 0,
    dados.status,
    dados.observacoes
  ]);

  return "Ordem de Serviço " + numeroOS + " cadastrada com sucesso!";
}
