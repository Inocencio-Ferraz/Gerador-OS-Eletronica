const ID_PLANILHA = "1X0jjZKfoP7WqIWlVQ4mKjq8aFBkOf-Ex-0IzNdMKfF0";

const STATUS_VALIDOS = [
  "Recebido",
  "Em análise",
  "Aguardando aprovação",
  "Aguardando peças",
  "Em manutenção",
  "Concluído",
  "Entregue",
  "Cancelado"
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Controle de OS - Eletrônica")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function obterAbas() {
  const planilha = SpreadsheetApp.openById(ID_PLANILHA);
  const clientes = planilha.getSheetByName("Clientes");
  const ordens = planilha.getSheetByName("Ordens_Servico");

  if (!clientes || !ordens) {
    throw new Error("As abas Clientes e Ordens_Servico precisam existir.");
  }

  return { clientes, ordens };
}

function numero(valor) {
  if (typeof valor === "number") return valor;
  return Number(String(valor || "0").replace(",", ".")) || 0;
}

function telefoneNormalizado(valor) {
  return String(valor || "").replace(/\D/g, "");
}

function formatarData(valor) {
  if (!valor) return "";
  const data = valor instanceof Date ? valor : new Date(valor);
  if (isNaN(data.getTime())) return String(valor);

  return Utilities.formatDate(
    data,
    Session.getScriptTimeZone(),
    "dd/MM/yyyy HH:mm"
  );
}

function formatarDataISO(valor) {
  if (!valor) return "";
  const data = valor instanceof Date ? valor : new Date(valor);
  if (isNaN(data.getTime())) return "";

  return Utilities.formatDate(
    data,
    Session.getScriptTimeZone(),
    "yyyy-MM-dd"
  );
}

function localizarLinhaOS(aba, numeroOS) {
  const ultima = aba.getLastRow();
  if (ultima < 2) throw new Error("Nenhuma OS cadastrada.");

  const ids = aba.getRange(2, 1, ultima - 1, 1).getValues();

  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(numeroOS)) {
      return i + 2;
    }
  }

  throw new Error("Ordem de Serviço não encontrada.");
}

function salvarOS(dados) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const { clientes, ordens } = obterAbas();

    const nome = String(dados.nome || "").trim();
    const telefone = String(dados.telefone || "").trim();
    const marca = String(dados.marca || "").trim();
    const valor = numero(dados.valor);
    const status = String(dados.status || "Recebido");

    if (!nome || !telefoneNormalizado(telefone)) {
      throw new Error("Preencha o nome e o telefone do cliente.");
    }

    if (valor < 0 || !isFinite(valor)) {
      throw new Error("O valor informado é inválido.");
    }

    if (!STATUS_VALIDOS.includes(status)) {
      throw new Error("Status inválido.");
    }

    const agora = new Date();
    const telefoneBusca = telefoneNormalizado(telefone);
    let idCliente = "";

    if (clientes.getLastRow() >= 2) {
      const registros = clientes.getRange(
        2, 1, clientes.getLastRow() - 1, 4
      ).getValues();

      for (const linha of registros) {
        if (telefoneNormalizado(linha[2]) === telefoneBusca) {
          idCliente = String(linha[0]);
          break;
        }
      }
    }

    if (!idCliente) {
      idCliente = "CLI" + Utilities.getUuid()
        .slice(0, 8).toUpperCase();

      clientes.appendRow([idCliente, nome, telefone, agora]);
    }

    const numeroOS = "OS" +
      Utilities.formatDate(
        agora,
        Session.getScriptTimeZone(),
        "yyyyMMdd"
      ) +
      Utilities.getUuid().slice(0, 6).toUpperCase();

    ordens.appendRow([
      numeroOS,
      idCliente,
      agora,
      marca,
      dados.polegadas || "",
      dados.diagnostico || "",
      dados.servico || "",
      valor,
      status,
      dados.observacoes || "",
      0,
      "",
      ""
    ]);

    return "Ordem de Serviço " + numeroOS + " cadastrada com sucesso!";
  } finally {
    lock.releaseLock();
  }
}

function listarOS() {
  const { clientes, ordens } = obterAbas();
  const mapaClientes = {};

  if (clientes.getLastRow() >= 2) {
    const dadosClientes = clientes.getRange(
      2, 1, clientes.getLastRow() - 1, 4
    ).getValues();

    dadosClientes.forEach(linha => {
      const id = String(linha[0] || "");
      if (id) {
        mapaClientes[id] = {
          nome: String(linha[1] || ""),
          telefone: String(linha[2] || "")
        };
      }
    });
  }

  if (ordens.getLastRow() < 2) return [];

  const registros = ordens.getRange(
    2, 1, ordens.getLastRow() - 1, 13
  ).getValues();

  return registros
    .filter(linha => String(linha[0] || "").startsWith("OS"))
    .map(linha => {
      const cliente = mapaClientes[String(linha[1])] || {
        nome: "Cliente não localizado",
        telefone: ""
      };

      return {
        numero: String(linha[0]),
        idCliente: String(linha[1]),
        data: formatarData(linha[2]),
        dataISO: formatarDataISO(linha[2]),
        nome: cliente.nome,
        telefone: cliente.telefone,
        marca: String(linha[3] || ""),
        polegadas: String(linha[4] || ""),
        diagnostico: String(linha[5] || ""),
        servico: String(linha[6] || ""),
        valor: numero(linha[7]),
        status: String(linha[8] || "Recebido"),
        observacoes: String(linha[9] || ""),
        recebido: numero(linha[10]),
        formaPagamento: String(linha[11] || ""),
        dataPagamento: formatarData(linha[12])
      };
    })
    .reverse();
}

function atualizarStatusOS(numeroOS, novoStatus) {
  if (!STATUS_VALIDOS.includes(novoStatus)) {
    throw new Error("Status inválido.");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const { ordens } = obterAbas();
    const linha = localizarLinhaOS(ordens, numeroOS);
    ordens.getRange(linha, 9).setValue(novoStatus);
    return "Status atualizado com sucesso!";
  } finally {
    lock.releaseLock();
  }
}

function excluirOS(numeroOS) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const { ordens } = obterAbas();
    const linha = localizarLinhaOS(ordens, numeroOS);
    ordens.deleteRow(linha);
    return "Ordem de Serviço excluída com sucesso.";
  } finally {
    lock.releaseLock();
  }
}

function registrarPagamento(numeroOS, valorPagamento, formaPagamento) {
  const valor = numero(valorPagamento);

  if (!isFinite(valor) || valor <= 0) {
    throw new Error("Informe um valor de pagamento válido.");
  }

  const formas = ["Dinheiro", "Pix", "Cartão", "Transferência"];
  if (!formas.includes(formaPagamento)) {
    throw new Error("Forma de pagamento inválida.");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const { ordens } = obterAbas();
    const linha = localizarLinhaOS(ordens, numeroOS);
    const dados = ordens.getRange(linha, 8, 1, 6).getValues()[0];

    const totalOS = numero(dados[0]);
    const recebidoAnterior = numero(dados[3]);
    const saldo = Math.round((totalOS - recebidoAnterior) * 100) / 100;

    if (valor > saldo + 0.001) {
      throw new Error("O pagamento não pode ser maior que o valor pendente.");
    }

    const novoTotal = Math.round((recebidoAnterior + valor) * 100) / 100;

    ordens.getRange(linha, 11, 1, 3).setValues([[
      novoTotal,
      formaPagamento,
      new Date()
    ]]);

    return "Pagamento registrado com sucesso!";
  } finally {
    lock.releaseLock();
  }
}
