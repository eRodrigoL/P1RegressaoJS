const xlsx = require("xlsx");

// Ler o arquivo .xlsx
const workbook = xlsx.readFile("005-Dataset.xlsx");

// Função da regressão linear
function regressaoLinear(pontos) {
  const n = pontos.length;

  // Somatórios
  let somaX = 0,
    somaY = 0,
    somaXY = 0,
    somaX2 = 0;

  for (let i = 0; i < n; i++) {
    const [x, y] = pontos[i];
    somaX += x;
    somaY += y;
    somaXY += x * y;
    somaX2 += x * x;
  }

  // Coeficiente angular (m)
  const m = (n * somaXY - somaX * somaY) / (n * somaX2 - somaX * somaX);

  // Coeficiente linear (b)
  const b = (somaY - m * somaX) / n;

  // Calculo de R2
  const meanY = pontos.reduce((sum, [_, y]) => sum + y, 0) / n;

  // Soma dos quadrados totais (SST)
  const sst = pontos.reduce((sum, [_, y]) => sum + Math.pow(y - meanY, 2), 0);

  // Soma dos quadrados dos resíduos (SSE)
  const sse = pontos.reduce((sum, [x, y]) => {
    const yPred = m * x + b;
    return sum + Math.pow(y - yPred, 2);
  }, 0);

  // Coeficiente de determinação R^2
  const r2 = 1 - sse / sst;

  // Arredondar os resultados para 3 casas decimais
  return {
    m: parseFloat(m.toFixed(3)),
    b: parseFloat(b.toFixed(3)),
    r2: parseFloat(r2.toFixed(3)),
  };
}

// Armazenar os resultados
const resultados = [];

// Iterar sobre todas as planilhas do arquivo
workbook.SheetNames.forEach((sheetName) => {
  const amostra = workbook.Sheets[sheetName];

  // Converter a planilha para um array de objetos, omitindo o cabeçalho
  const dados = xlsx.utils.sheet_to_json(amostra, { header: 1 });

  // Remover o cabeçalho (primeira linha)
  dados.shift();

  // Converter os dados para o formato necessário (array de pontos) e filtrar pontos inválidos
  const pontos = dados
    .map((linha) => [linha[0], linha[1]]) // Converter para o formato de ponto
    .filter((ponto) => ponto[0] !== undefined && ponto[1] !== undefined);

  // Calcular a regressão linear
  const resultado = regressaoLinear(pontos);

  // Adicionar o resultado à lista, incluindo o nome da planilha
  resultados.push({
    planilha: sheetName,
    m: resultado.m,
    b: resultado.b,
    r2: resultado.r2,
  });
});

// Calcular a média dos coeficientes
const media = {
  planilha: "Média",
  m: (
    resultados.reduce((sum, resultado) => sum + resultado.m, 0) /
    resultados.length
  ).toFixed(3),
  b: (
    resultados.reduce((sum, resultado) => sum + resultado.b, 0) /
    resultados.length
  ).toFixed(3),
  r2: (
    resultados.reduce((sum, resultado) => sum + resultado.r2, 0) /
    resultados.length
  ).toFixed(3),
};

// Exibir a tabela de resultados
console.table(resultados);

// Adicionar e exibir a linha de média
console.table([media]);
