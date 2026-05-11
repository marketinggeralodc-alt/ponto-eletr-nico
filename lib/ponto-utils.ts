import * as XLSX from "xlsx"

export interface DadosPonto {
  nome: string
  dia: number
  mes: number
  ano: number
  entrada1: string
  saida1: string
  entrada2: string
  saida2: string
  status: string
}

export interface ResultadoProcessamento {
  dados: DadosPonto[]
  funcionarios: string[]
  mesAno: string
  erros: string[]
  debug: string[]
}

// Normaliza nome para comparação
export function normalizarNome(nome: string): string {
  return nome
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9\s]/g, "")
    .trim()
}

// Extrai todos os horários de uma string/célula
function extrairHorarios(val: unknown): string[] {
  if (val === null || val === undefined) return []
  const str = String(val)
  const matches = str.match(/\d{1,2}:\d{2}/g)
  return matches || []
}

// Converte horário HH:MM para minutos
function horarioParaMinutos(horario: string): number {
  const match = horario.match(/(\d{1,2}):(\d{2})/)
  if (!match) return -1
  return parseInt(match[1]) * 60 + parseInt(match[2])
}

// Limpa e ordena horários, removendo duplicatas próximas
function processarHorarios(horarios: string[]): string[] {
  if (horarios.length === 0) return []

  // Converte para minutos e ordena
  const horariosMinutos = horarios
    .map((h) => ({ original: h, minutos: horarioParaMinutos(h) }))
    .filter((h) => h.minutos >= 0)
    .sort((a, b) => a.minutos - b.minutos)

  if (horariosMinutos.length === 0) return []

  // Remove duplicatas muito próximas (menos de 30 min)
  const horariosLimpos: typeof horariosMinutos = [horariosMinutos[0]]

  for (let i = 1; i < horariosMinutos.length; i++) {
    const atual = horariosMinutos[i]
    const anterior = horariosLimpos[horariosLimpos.length - 1]

    if (atual.minutos - anterior.minutos >= 30) {
      horariosLimpos.push(atual)
    }
  }

  return horariosLimpos.map((h) => h.original)
}

// Normaliza horário para formato HH:MM
function normalizarHorario(horario: string): string {
  const match = horario.match(/(\d{1,2}):(\d{2})/)
  if (!match) return horario
  return `${match[1].padStart(2, "0")}:${match[2]}`
}

// Cria registro de ponto com os horários
function criarRegistro(
  nome: string,
  dia: number,
  mes: number,
  ano: number,
  horarios: string[],
  isSabado: boolean,
  isDomingo: boolean
): DadosPonto {
  const qtd = horarios.length
  let entrada1 = ""
  let saida1 = ""
  let entrada2 = ""
  let saida2 = ""
  let status = ""

  // Normaliza horários
  const horariosNorm = horarios.map(normalizarHorario)

  if (isDomingo) {
    if (qtd >= 2) {
      status = `Domingo (${qtd} batidas)`
      entrada1 = horariosNorm[0]
      saida2 = horariosNorm[qtd - 1]
    } else if (qtd === 1) {
      status = `Domingo (${qtd} batida)`
      entrada1 = horariosNorm[0]
    }
  } else if (isSabado) {
    if (qtd >= 2) {
      status = "OK"
      entrada1 = horariosNorm[0]
      saida2 = horariosNorm[qtd - 1]
    } else if (qtd === 1) {
      status = `Incompleto (${qtd} batida)`
      entrada1 = horariosNorm[0]
    }
  } else {
    if (qtd >= 4) {
      status = qtd === 4 ? "OK" : `OK (${qtd} batidas)`
      entrada1 = horariosNorm[0]
      saida1 = horariosNorm[1]
      entrada2 = horariosNorm[2]
      saida2 = horariosNorm[qtd - 1]
    } else if (qtd === 3) {
      status = `Incompleto (${qtd} batidas)`
      entrada1 = horariosNorm[0]
      saida1 = horariosNorm[1]
      entrada2 = horariosNorm[2]
    } else if (qtd === 2) {
      status = `Incompleto (${qtd} batidas)`
      entrada1 = horariosNorm[0]
      saida2 = horariosNorm[1]
    } else if (qtd === 1) {
      status = `Incompleto (${qtd} batida)`
      entrada1 = horariosNorm[0]
    }
  }

  return { nome, dia, mes, ano, entrada1, saida1, entrada2, saida2, status }
}

// Detecta estrutura horizontal (dias nas colunas) - Layout: ID/Nome na linha, dias 1-31 nas colunas
function tentarParseHorizontal(
  linhas: unknown[][],
  debug: string[]
): { dados: DadosPonto[]; ano: number; mes: number } | null {
  debug.push("Tentando parse horizontal (dias nas colunas)")

  let anoPonto = new Date().getFullYear()
  let mesPonto = new Date().getMonth() + 1

  // Procura data no arquivo
  for (let i = 0; i < Math.min(10, linhas.length); i++) {
    const linhaTexto = linhas[i]?.map((x) => String(x)).join(" ") || ""
    const matchDate = linhaTexto.match(/(\d{4})-(\d{2})-\d{2}/)
    if (matchDate) {
      anoPonto = parseInt(matchDate[1])
      mesPonto = parseInt(matchDate[2])
      debug.push(`Periodo encontrado: ${mesPonto}/${anoPonto}`)
      break
    }
  }

  // Encontra linha com números de dias (1, 2, 3... 24 ou mais)
  let linhaDias = -1
  const diasColunas: Record<number, number> = {}

  for (let i = 0; i < linhas.length; i++) {
    const linha = linhas[i]
    if (!linha) continue

    let countDias = 0
    const diasEncontrados: number[] = []

    for (let j = 0; j < linha.length; j++) {
      const val = String(linha[j] || "").trim()
      if (/^[1-9]$|^[12][0-9]$|^3[01]$/.test(val)) {
        const dia = parseInt(val)
        if (!diasEncontrados.includes(dia)) {
          countDias++
          diasEncontrados.push(dia)
        }
      }
    }

    if (countDias >= 15) {
      linhaDias = i
      debug.push(`Linha de dias encontrada: linha ${i}, com ${countDias} dias`)

      for (let j = 0; j < linha.length; j++) {
        const val = String(linha[j] || "").trim()
        if (/^[1-9]$|^[12][0-9]$|^3[01]$/.test(val)) {
          const dia = parseInt(val)
          if (!diasColunas[dia]) {
            diasColunas[dia] = j
          }
        }
      }
      break
    }
  }

  if (linhaDias === -1) {
    debug.push("Nao encontrou linha de dias")
    return null
  }

  // Processa funcionários - procura padrões como "ID : N Nome : fulano" ou "Nome : fulano"
  interface FuncionarioBloco {
    nome: string
    linhaInicio: number
    linhaFim: number
  }

  const blocosFuncionarios: FuncionarioBloco[] = []

  for (let i = 0; i < linhas.length; i++) {
    const linha = linhas[i]
    if (!linha) continue

    const linhaTexto = linha.map((x) => String(x || "").toLowerCase()).join(" ")

    if (
      (linhaTexto.includes("id") && linhaTexto.includes("nome")) ||
      linhaTexto.includes("nome :") ||
      linhaTexto.includes("nome:")
    ) {
      let nomeFuncionario = ""

      // Procura célula após "nome" ou "nome :"
      for (let j = 0; j < linha.length; j++) {
        const val = String(linha[j] || "").toLowerCase().trim()
        if (val === "nome" || val === "nome :" || val === "nome:") {
          for (let k = j + 1; k < Math.min(j + 5, linha.length); k++) {
            const candidato = String(linha[k] || "").trim()
            if (
              candidato &&
              candidato !== ":" &&
              candidato.length > 1 &&
              !candidato.toLowerCase().includes("dept") &&
              !candidato.toLowerCase().includes("id")
            ) {
              nomeFuncionario = candidato
              break
            }
          }
          break
        }
      }

      // Regex na linha inteira
      if (!nomeFuncionario) {
        const match = linhaTexto.match(/nome\s*:?\s*([a-záéíóúâêîôûãõç\s]+)/i)
        if (match && match[1]) {
          const nome = match[1].trim()
          if (nome.length > 1 && !nome.includes("dept")) {
            nomeFuncionario = nome
          }
        }
      }

      if (nomeFuncionario) {
        debug.push(`Funcionario: ${nomeFuncionario} (linha ${i})`)

        if (blocosFuncionarios.length > 0) {
          blocosFuncionarios[blocosFuncionarios.length - 1].linhaFim = i - 1
        }

        blocosFuncionarios.push({
          nome: nomeFuncionario,
          linhaInicio: i + 1,
          linhaFim: linhas.length - 1,
        })
      }
    }
  }

  if (blocosFuncionarios.length === 0) {
    debug.push("Nenhum funcionario encontrado")
    return null
  }

  debug.push(`Total de funcionarios: ${blocosFuncionarios.length}`)

  // Coleta horários de cada funcionário
  const dados: DadosPonto[] = []

  for (const bloco of blocosFuncionarios) {
    const horariosPorDia: Record<number, string[]> = {}

    for (let row = bloco.linhaInicio; row <= Math.min(bloco.linhaFim, bloco.linhaInicio + 10); row++) {
      const linha = linhas[row]
      if (!linha) continue

      const linhaTexto = linha.map((x) => String(x || "").toLowerCase()).join(" ")
      if (linhaTexto.includes("id") && linhaTexto.includes("nome")) {
        break
      }

      for (const [diaStr, colIdx] of Object.entries(diasColunas)) {
        const dia = parseInt(diaStr)
        if (colIdx >= linha.length) continue

        const horarios = extrairHorarios(linha[colIdx])
        if (horarios.length > 0) {
          if (!horariosPorDia[dia]) {
            horariosPorDia[dia] = []
          }
          horariosPorDia[dia].push(...horarios)
        }
      }
    }

    for (const [diaStr, horarios] of Object.entries(horariosPorDia)) {
      const dia = parseInt(diaStr)
      const horariosLimpos = processarHorarios(horarios)

      if (horariosLimpos.length === 0) continue

      const dataAtual = new Date(anoPonto, mesPonto - 1, dia)
      const isSabado = dataAtual.getDay() === 6
      const isDomingo = dataAtual.getDay() === 0

      const registro = criarRegistro(bloco.nome, dia, mesPonto, anoPonto, horariosLimpos, isSabado, isDomingo)
      dados.push(registro)
    }
  }

  debug.push(`Total de registros: ${dados.length}`)
  return { dados, ano: anoPonto, mes: mesPonto }
}

// Processa o arquivo do relógio
export function processarRelogio(workbook: XLSX.WorkBook): ResultadoProcessamento {
  const debug: string[] = []
  const erros: string[] = []

  debug.push(`Abas encontradas: ${workbook.SheetNames.join(", ")}`)

  let dadosTotais: DadosPonto[] = []
  let anoFinal = new Date().getFullYear()
  let mesFinal = new Date().getMonth() + 1

  for (const sheetName of workbook.SheetNames) {
    debug.push(`--- Processando aba: ${sheetName} ---`)

    const sheet = workbook.Sheets[sheetName]
    const linhas: unknown[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" })

    debug.push(`Total de linhas: ${linhas.length}`)

    if (linhas.length === 0) {
      debug.push(`Aba vazia, pulando`)
      continue
    }

    // Mostra algumas linhas para debug
    for (let i = 0; i < Math.min(8, linhas.length); i++) {
      const linhaStr = linhas[i]?.slice(0, 12).map((x) => String(x || "").substring(0, 20)).join(" | ")
      debug.push(`L${i}: ${linhaStr}`)
    }

    const resultado = tentarParseHorizontal(linhas, debug)

    if (resultado && resultado.dados.length > 0) {
      debug.push(`Sucesso! ${resultado.dados.length} registros encontrados`)
      dadosTotais.push(...resultado.dados)
      anoFinal = resultado.ano
      mesFinal = resultado.mes
    } else {
      debug.push(`Nenhum registro encontrado nesta aba`)
    }
  }

  if (dadosTotais.length === 0) {
    erros.push("Estrutura do arquivo do relógio não reconhecida.")
    erros.push("Nenhum horário foi encontrado no arquivo do relógio. Verifique se é o arquivo correto.")
  }

  const funcionarios = [...new Set(dadosTotais.map((d) => d.nome))]

  return {
    dados: dadosTotais,
    funcionarios,
    mesAno: `${String(mesFinal).padStart(2, "0")}/${anoFinal}`,
    erros,
    debug,
  }
}

// Gera tabela de dados simples para usar com PROCV no Excel
export function gerarTabelaDados(
  dados: DadosPonto[],
  mesAno: string
): XLSX.WorkBook {
  // Ordena por nome e dia
  const dadosOrdenados = [...dados].sort((a, b) => {
    const nomeCompare = a.nome.localeCompare(b.nome)
    if (nomeCompare !== 0) return nomeCompare
    return a.dia - b.dia
  })

  // Cria planilha com cabeçalhos
  const cabecalhos = [
    "FUNCIONARIO",
    "DIA",
    "CHAVE_PROCV",
    "ENTRADA_1",
    "SAIDA_1",
    "ENTRADA_2",
    "SAIDA_2",
    "STATUS"
  ]

  const linhas: (string | number)[][] = [cabecalhos]

  for (const d of dadosOrdenados) {
    // Chave para PROCV: NOME_DIA (ex: "WYNER_7")
    const chaveProcv = `${d.nome.toUpperCase()}_${d.dia}`
    
    linhas.push([
      d.nome.toUpperCase(),
      d.dia,
      chaveProcv,
      d.entrada1,
      d.saida1,
      d.entrada2,
      d.saida2,
      d.status
    ])
  }

  // Cria workbook
  const ws = XLSX.utils.aoa_to_sheet(linhas)

  // Define largura das colunas
  ws["!cols"] = [
    { wch: 20 }, // FUNCIONARIO
    { wch: 6 },  // DIA
    { wch: 25 }, // CHAVE_PROCV
    { wch: 10 }, // ENTRADA_1
    { wch: 10 }, // SAIDA_1
    { wch: 10 }, // ENTRADA_2
    { wch: 10 }, // SAIDA_2
    { wch: 20 }, // STATUS
  ]

  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, ws, "Dados_Ponto")

  // Adiciona aba com instruções
  const instrucoes = [
    ["COMO USAR ESTA TABELA NO EXCEL"],
    [""],
    ["Esta tabela contém os dados extraídos do relógio de ponto."],
    ["Use a fórmula PROCV para preencher sua folha de ponto."],
    [""],
    ["PASSO A PASSO:"],
    [""],
    ["1. Abra sua folha de ponto original"],
    ["2. Cole esta tabela em uma aba separada (ex: 'Dados')"],
    ["3. Na sua folha de ponto, use estas fórmulas:"],
    [""],
    ["Para ENTRADA (coluna B):"],
    ['=PROCV($A6&"_"&DIA($A6);Dados!$C:$H;2;FALSO)'],
    [""],
    ["Para SAÍDA ALMOÇO (coluna C):"],
    ['=PROCV($A6&"_"&DIA($A6);Dados!$C:$H;3;FALSO)'],
    [""],
    ["Para ENTRADA TARDE (coluna E):"],
    ['=PROCV($A6&"_"&DIA($A6);Dados!$C:$H;4;FALSO)'],
    [""],
    ["Para SAÍDA (coluna F):"],
    ['=PROCV($A6&"_"&DIA($A6);Dados!$C:$H;5;FALSO)'],
    [""],
    ["DICA: A coluna CHAVE_PROCV combina NOME_DIA."],
    ["Ajuste a fórmula conforme o nome da aba do funcionário."],
    [""],
    ["Exemplo para WYNER no dia 7:"],
    ['=PROCV("WYNER_7";Dados!$C:$H;2;FALSO)'],
  ]

  const wsInstrucoes = XLSX.utils.aoa_to_sheet(instrucoes)
  wsInstrucoes["!cols"] = [{ wch: 60 }]
  XLSX.utils.book_append_sheet(wb, wsInstrucoes, "Instrucoes")

  // Adiciona aba por funcionário para facilitar
  const funcionarios = [...new Set(dados.map(d => d.nome))]
  
  for (const func of funcionarios) {
    const dadosFunc = dadosOrdenados.filter(d => d.nome === func)
    const linhasFunc: (string | number)[][] = [
      ["DIA", "ENTRADA_1", "SAIDA_1", "ENTRADA_2", "SAIDA_2", "STATUS"]
    ]
    
    for (const d of dadosFunc) {
      linhasFunc.push([
        d.dia,
        d.entrada1,
        d.saida1,
        d.entrada2,
        d.saida2,
        d.status
      ])
    }
    
    const wsFunc = XLSX.utils.aoa_to_sheet(linhasFunc)
    wsFunc["!cols"] = [
      { wch: 6 },  // DIA
      { wch: 10 }, // ENTRADA_1
      { wch: 10 }, // SAIDA_1
      { wch: 10 }, // ENTRADA_2
      { wch: 10 }, // SAIDA_2
      { wch: 20 }, // STATUS
    ]
    
    // Limita nome da aba a 31 caracteres (limite do Excel)
    const nomeAba = func.substring(0, 31).toUpperCase()
    XLSX.utils.book_append_sheet(wb, wsFunc, nomeAba)
  }

  return wb
}
