const xlsx = require('xlsx');
const path = require('path');

// Caminho para o arquivo XLSX
const filePath = path.join(__dirname, 'cotas1730912864850.xlsx');

// Ler o arquivo XLSX
const workbook = xlsx.readFile(filePath);
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// Converter a data serial para o tipo Date em JavaScript
function excelDateToJSDate(serial) {
    const excelStartDate = new Date(1900, 0, 1);
    const daysOffset = serial - 2;
    excelStartDate.setDate(excelStartDate.getDate() + daysOffset);
    return excelStartDate;
}

// Função para formatar a data no formato dd/mm/YYYY
function formatDate(date) {
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    return `${day}/${month}/${year}`;
}

// Converter a planilha para JSON e extrair as colunas "Data" e "Cota"
const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

const cotaData = data.slice(1).map(row => ({
    date: excelDateToJSDate(row[2]), // Converte a data serial para o tipo Date
    cota: parseFloat(row[3])         // Supondo que a coluna "Cota" é a quarta coluna
})).filter(entry => !isNaN(entry.cota)); // Filtra valores não numéricos

// Ordenar os dados por data
cotaData.sort((a, b) => new Date(a.date) - new Date(b.date));

// Extrair apenas os valores de cota
const cotaValues = cotaData.map(entry => entry.cota);

// Dividir em 80% para Treinamento e 20% para Validação
const splitIndex = Math.floor(cotaValues.length * 0.8);
const trainData = cotaValues.slice(0, splitIndex);
const validationData = cotaValues.slice(splitIndex);

// Função para calcular a Média Móvel Exponencial (EMA)
function exponentialMovingAverage(data, alpha = 0.1) {
    let ema = [data[0]];
    for (let i = 1; i < data.length; i++) {
        ema.push(alpha * data[i] + (1 - alpha) * ema[i - 1]);
    }
    return ema;
}

// Aplicar EMA ao conjunto de treinamento e obter as previsões completas
const predictions = exponentialMovingAverage(cotaValues, 0.1);

// Função para calcular RMSE (Erro Médio Quadrático da Raiz)
function calculateRMSE(predictions, validationData) {
    const errors = validationData.map((value, index) => {
        const predictionIndex = splitIndex + index;
        return (value - predictions[predictionIndex]) ** 2;
    });
    const meanError = errors.reduce((sum, value) => sum + value, 0) / validationData.length;
    return Math.sqrt(meanError);
}

// Calcular o RMSE e o MAE para as previsões
const rmse = calculateRMSE(predictions, validationData);

console.log("RMSE (Erro Médio Quadrático das previsões):", rmse);

// Exibir as previsões com datas para o conjunto de validação
const validationPredictions = validationData.map((value, index) => ({
    date: formatDate(cotaData[splitIndex + index].date),
    predictedCota: predictions[splitIndex + index]
}));
console.log("Previsões para o Conjunto de Validação:", validationPredictions.reverse());

