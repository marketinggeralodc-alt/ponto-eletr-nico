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

// Lista de feriados nacionais brasileiros (fixos)
// Formato: "DD/MM"
const FERIADOS_FIXOS = [
  "01/01", // Confraternização Universal
  "21/04", // Tiradentes
  "01/05", // Dia do Trabalho
  "07/09", // Independência do Brasil
  "12/10", // Nossa Senhora Aparecida
  "02/11", // Finados
  "15/11", // Proclamação da República
  "25/12", // Natal
]

// Feriados móveis para anos específicos (Carnaval, Sexta-feira Santa, Corpus Christi)
// Precisa ser atualizado manualmente ou calculado
const FERIADOS_MOVEIS: Record<number, string[]> = {
  2024: [
    "12/02", "13/02", // Carnaval
    "29/03", // Sexta-feira Santa
    "30/05", // Corpus Christi
  ],
  2025: [
    "03/03", "04/03", // Carnaval
    "18/04", // Sexta-feira Santa
    "19/06", // Corpus Christi
  ],
  2026: [
    "16/02", "17/02", // Carnaval
    "03/04", // Sexta-feira Santa
    "04/06", // Corpus Christi
  ],
}

// Verifica se uma data é feriado
export function verificarFeriado(dia: number, mes: number, ano: number): { isFeriado: boolean; nome?: string } {
  const dataStr = `${String(dia).padStart(2, "0")}/${String(mes).padStart(2, "0")}`
  
  // Verifica feriados fixos
  const feriadosFixosNomes: Record<string, string> = {
    "01/01": "Confraternização Universal",
    "21/04": "Tiradentes",
    "01/05": "Dia do Trabalho",
    "07/09": "Independência do Brasil",
    "12/10": "Nossa Senhora Aparecida",
    "02/11": "Finados",
    "15/11": "Proclamação da República",
    "25/12": "Natal",
  }
  
  if (FERIADOS_FIXOS.includes(dataStr)) {
    return { isFeriado: true, nome: feriadosFixosNomes[dataStr] }
  }
  
  // Verifica feriados móveis do ano
  const feriadosMoveis = FERIADOS_MOVEIS[ano] || []
  if (feriadosMoveis.includes(dataStr)) {
    return { isFeriado: true, nome: "Feriado Móvel" }
  }
  
  return { isFeriado: false }
}

// Calcula duração do almoço em minutos
export function calcularDuracaoAlmoco(saida1: string, entrada2: string): number {
  if (!saida1 || !entrada2) return 0
  
  const minutosSaida = horarioParaMinutos(saida1)
  const minutosEntrada = horarioParaMinutos(entrada2)
  
  if (minutosSaida < 0 || minutosEntrada < 0) return 0
  
  return minutosEntrada - minutosSaida
}

// Calcula total de horas trabalhadas em minutos
export function calcularHorasTrabalhadas(entrada1: string, saida1: string, entrada2: string, saida2: string): number {
  let total = 0
  
  if (entrada1 && saida1) {
    const min1 = horarioParaMinutos(entrada1)
    const min2 = horarioParaMinutos(saida1)
    if (min1 >= 0 && min2 >= 0) {
      total += min2 - min1
    }
  }
  
  if (entrada2 && saida2) {
    const min3 = horarioParaMinutos(entrada2)
    const min4 = horarioParaMinutos(saida2)
    if (min3 >= 0 && min4 >= 0) {
      total += min4 - min3
    }
  }
  
  // Se só tem entrada1 e saida2 (sem almoço registrado)
  if (entrada1 && saida2 && !saida1 && !entrada2) {
    const min1 = horarioParaMinutos(entrada1)
    const min4 = horarioParaMinutos(saida2)
    if (min1 >= 0 && min4 >= 0) {
      total = min4 - min1
    }
  }
  
  return total
}

// Formata minutos para HH:MM (suporta valores negativos)
export function minutosParaHorario(minutos: number): string {
  const horas = Math.floor(Math.abs(minutos) / 60)
  const mins = Math.abs(minutos) % 60
  const sinal = minutos < 0 ? "-" : ""
  return `${sinal}${String(horas).padStart(2, "0")}:${String(mins).padStart(2, "0")}`
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

// Gera tabela de dados completa com cálculos de horas extras
export function gerarTabelaDados(
  dados: DadosPonto[],
  mesAno: string
): XLSX.WorkBook {
  const wb = XLSX.utils.book_new()
  
  // Extrai mês e ano do mesAno (formato "MM/YYYY")
  const [mesStr, anoStr] = mesAno.split("/")
  const mes = parseInt(mesStr) || new Date().getMonth() + 1
  const ano = parseInt(anoStr) || new Date().getFullYear()
  
  // Obtém o número de dias no mês
  const diasNoMes = new Date(ano, mes, 0).getDate()
  
  // Agrupa dados por funcionário
  const funcionarios = [...new Set(dados.map(d => d.nome))]
  
  // Cria uma aba completa para cada funcionário
  for (const func of funcionarios) {
    const dadosFunc = dados.filter(d => d.nome === func)
    const dadosPorDia: Record<number, DadosPonto> = {}
    for (const d of dadosFunc) {
      dadosPorDia[d.dia] = d
    }
    
    // Cria estrutura da planilha do funcionário
    const linhas: (string | number)[][] = []
    
    // Cabeçalho com informações do funcionário
    linhas.push(["FOLHA DE PONTO - " + mesAno])
    linhas.push([])
    linhas.push(["NOME:", func.toUpperCase(), "", "FUNÇÃO:", ""])
    linhas.push([])
    
    // Cabeçalho da tabela (linha 5 - índice 4)
    linhas.push([
      "DIA",
      "DIA SEMANA",
      "ENTRADA 1",
      "SAÍDA ALMOÇO",
      "ENTRADA 2",
      "SAÍDA",
      "TOTAL HORAS",
      "HORAS EXTRAS",
      "FERIADO/OBS",
      "ALERTA"
    ])
    
    // Jornada padrão em minutos (8 horas)
    const JORNADA_PADRAO = 8 * 60
    const ALMOCO_MINIMO = 60 // 1 hora
    
    let totalHorasExtras = 0
    const alertas: string[] = []
    const observacoesSistema: string[] = []
    
    // Preenche cada dia do mês
    for (let dia = 1; dia <= diasNoMes; dia++) {
      const dataAtual = new Date(ano, mes - 1, dia)
      const diaSemana = dataAtual.getDay()
      const diasSemanaTexto = ["DOM", "SEG", "TER", "QUA", "QUI", "SEX", "SAB"]
      const diaSemanaStr = diasSemanaTexto[diaSemana]
      
      const registro = dadosPorDia[dia]
      const feriado = verificarFeriado(dia, mes, ano)
      const isDomingo = diaSemana === 0
      const isSabado = diaSemana === 6
      
      let entrada1 = ""
      let saida1 = ""
      let entrada2 = ""
      let saida2 = ""
      let totalHoras = ""
      let horasExtras = ""
      let feriadoObs = ""
      let alerta = ""
      
      if (registro) {
        entrada1 = registro.entrada1
        saida1 = registro.saida1
        entrada2 = registro.entrada2
        saida2 = registro.saida2
        
        // Calcula total de horas trabalhadas
        const minutosTrabalhados = calcularHorasTrabalhadas(entrada1, saida1, entrada2, saida2)
        
        if (minutosTrabalhados > 0) {
          totalHoras = minutosParaHorario(minutosTrabalhados)
          
          // Calcula horas extras
          if (isDomingo || isSabado || feriado.isFeriado) {
            // Fim de semana ou feriado: todas as horas são extras
            if (minutosTrabalhados > 0) {
              horasExtras = minutosParaHorario(minutosTrabalhados)
              totalHorasExtras += minutosTrabalhados
            }
          } else {
            // Dia normal: horas extras são as que excedem a jornada
            const extras = minutosTrabalhados - JORNADA_PADRAO
            if (extras > 0) {
              horasExtras = minutosParaHorario(extras)
              totalHorasExtras += extras
            } else if (extras < 0) {
              // Horas faltantes (negativo)
              horasExtras = minutosParaHorario(extras)
            }
          }
        }
        
        // Verifica duração do almoço
        const duracaoAlmoco = calcularDuracaoAlmoco(saida1, entrada2)
        if (duracaoAlmoco > 0 && duracaoAlmoco < ALMOCO_MINIMO) {
          alerta = `ALMOÇO < 1h (${minutosParaHorario(duracaoAlmoco)})`
          alertas.push(`Dia ${dia}: Almoço de apenas ${minutosParaHorario(duracaoAlmoco)}`)
        }
        
        // Verifica registro incompleto
        if (registro.status.includes("Incompleto")) {
          if (alerta) alerta += " | "
          alerta += registro.status
          observacoesSistema.push(`Dia ${dia}: ${registro.status}`)
        }
      }
      
      // Define observação de feriado/fim de semana
      if (feriado.isFeriado) {
        feriadoObs = `FERIADO: ${feriado.nome}`
      } else if (isDomingo) {
        feriadoObs = "DOMINGO"
      } else if (isSabado) {
        feriadoObs = "SÁBADO"
      }
      
      linhas.push([
        dia,
        diaSemanaStr,
        entrada1,
        saida1,
        entrada2,
        saida2,
        totalHoras,
        horasExtras,
        feriadoObs,
        alerta
      ])
    }
    
    // Linha de totais (após os dias)
    linhas.push([])
    linhas.push(["", "", "", "", "", "TOTAL HORAS EXTRAS:", minutosParaHorario(totalHorasExtras), "", "", ""])
    
    // Seção de observações
    linhas.push([])
    linhas.push(["OBSERVAÇÕES DO SISTEMA:"])
    
    if (alertas.length > 0) {
      linhas.push(["ALERTAS DE ALMOÇO:"])
      for (const alerta of alertas) {
        linhas.push([alerta])
      }
    }
    
    if (observacoesSistema.length > 0) {
      linhas.push([])
      linhas.push(["REGISTROS INCOMPLETOS:"])
      for (const obs of observacoesSistema) {
        linhas.push([obs])
      }
    }
    
    // Espaço para observações manuais
    linhas.push([])
    linhas.push(["OUTRAS OBSERVAÇÕES:"])
    linhas.push([""])
    linhas.push([""])
    linhas.push([""])
    
    // Assinatura
    linhas.push([])
    linhas.push(["", "", "", "___________________________"])
    linhas.push(["", "", "", "Assinatura do Funcionário"])
    
    // Cria worksheet
    const ws = XLSX.utils.aoa_to_sheet(linhas)
    
    // Define largura das colunas
    ws["!cols"] = [
      { wch: 6 },   // DIA
      { wch: 10 },  // DIA SEMANA
      { wch: 12 },  // ENTRADA 1
      { wch: 14 },  // SAÍDA ALMOÇO
      { wch: 12 },  // ENTRADA 2
      { wch: 10 },  // SAÍDA
      { wch: 12 },  // TOTAL HORAS
      { wch: 12 },  // HORAS EXTRAS
      { wch: 25 },  // FERIADO/OBS
      { wch: 30 },  // ALERTA
    ]
    
    // Merge para o título
    ws["!merges"] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 9 } }, // Título
      { s: { r: 2, c: 1 }, e: { r: 2, c: 2 } }, // Nome
      { s: { r: 2, c: 4 }, e: { r: 2, c: 6 } }, // Função
    ]
    
    // Limita nome da aba a 31 caracteres
    const nomeAba = func.substring(0, 31).toUpperCase()
    XLSX.utils.book_append_sheet(wb, ws, nomeAba)
  }
  
  // Cria aba de resumo geral
  const resumoLinhas: (string | number)[][] = []
  resumoLinhas.push(["RESUMO DE HORAS EXTRAS - " + mesAno])
  resumoLinhas.push([])
  resumoLinhas.push(["FUNCIONÁRIO", "TOTAL HORAS EXTRAS", "OBSERVAÇÕES"])
  
  for (const func of funcionarios) {
    const dadosFunc = dados.filter(d => d.nome === func)
    let totalExtras = 0
    
    for (const d of dadosFunc) {
      const dataAtual = new Date(d.ano, d.mes - 1, d.dia)
      const diaSemana = dataAtual.getDay()
      const feriado = verificarFeriado(d.dia, d.mes, d.ano)
      
      const minutosTrabalhados = calcularHorasTrabalhadas(d.entrada1, d.saida1, d.entrada2, d.saida2)
      
      if (diaSemana === 0 || diaSemana === 6 || feriado.isFeriado) {
        totalExtras += minutosTrabalhados
      } else {
        const extras = minutosTrabalhados - (8 * 60)
        if (extras > 0) totalExtras += extras
      }
    }
    
    resumoLinhas.push([func.toUpperCase(), minutosParaHorario(totalExtras), ""])
  }
  
  const wsResumo = XLSX.utils.aoa_to_sheet(resumoLinhas)
  wsResumo["!cols"] = [
    { wch: 25 },  // FUNCIONÁRIO
    { wch: 20 },  // TOTAL HORAS EXTRAS
    { wch: 40 },  // OBSERVAÇÕES
  ]
  
  XLSX.utils.book_append_sheet(wb, wsResumo, "RESUMO")
  
  // Cria aba com dados brutos (para PROCV)
  const dadosOrdenados = [...dados].sort((a, b) => {
    const nomeCompare = a.nome.localeCompare(b.nome)
    if (nomeCompare !== 0) return nomeCompare
    return a.dia - b.dia
  })
  
  const cabecalhosDados = [
    "FUNCIONARIO",
    "DIA",
    "CHAVE_PROCV",
    "ENTRADA_1",
    "SAIDA_1",
    "ENTRADA_2",
    "SAIDA_2",
    "STATUS"
  ]
  
  const linhasDados: (string | number)[][] = [cabecalhosDados]
  
  for (const d of dadosOrdenados) {
    const chaveProcv = `${d.nome.toUpperCase()}_${d.dia}`
    linhasDados.push([
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
  
  const wsDados = XLSX.utils.aoa_to_sheet(linhasDados)
  wsDados["!cols"] = [
    { wch: 20 }, // FUNCIONARIO
    { wch: 6 },  // DIA
    { wch: 25 }, // CHAVE_PROCV
    { wch: 10 }, // ENTRADA_1
    { wch: 10 }, // SAIDA_1
    { wch: 10 }, // ENTRADA_2
    { wch: 10 }, // SAIDA_2
    { wch: 20 }, // STATUS
  ]
  
  XLSX.utils.book_append_sheet(wb, wsDados, "Dados_PROCV")
  
  // Aba de instruções
  const instrucoes = [
    ["INSTRUÇÕES DE USO"],
    [""],
    ["Esta planilha contém:"],
    [""],
    ["1. Uma aba para cada funcionário com:"],
    ["   - Espaço para nome e função"],
    ["   - Todos os dias do mês com horários"],
    ["   - Cálculo automático de horas trabalhadas"],
    ["   - Cálculo de horas extras (positivas e negativas)"],
    ["   - Identificação de feriados, sábados e domingos"],
    ["   - Alertas quando almoço for menor que 1 hora"],
    ["   - Total de horas extras do mês"],
    ["   - Espaço para observações"],
    [""],
    ["2. Aba RESUMO com total de horas extras de cada funcionário"],
    [""],
    ["3. Aba Dados_PROCV para usar com fórmulas PROCV"],
    [""],
    ["LEGENDA:"],
    ["- FERIADO: Dia é feriado nacional"],
    ["- SÁBADO/DOMINGO: Fim de semana"],
    ["- ALMOÇO < 1h: Intervalo de almoço menor que 1 hora"],
    ["- Incompleto: Faltam batidas no registro"],
    [""],
    ["OBSERVAÇÕES:"],
    ["- Horas extras em feriados/fins de semana = todas as horas trabalhadas"],
    ["- Horas extras em dias normais = horas além de 8h"],
    ["- Valores negativos indicam horas faltantes"],
    [""],
    ["FERIADOS CONSIDERADOS:"],
    ["- 01/01: Confraternização Universal"],
    ["- 21/04: Tiradentes"],
    ["- 01/05: Dia do Trabalho"],
    ["- 07/09: Independência do Brasil"],
    ["- 12/10: Nossa Senhora Aparecida"],
    ["- 02/11: Finados"],
    ["- 15/11: Proclamação da República"],
    ["- 25/12: Natal"],
    ["- Carnaval, Sexta-feira Santa, Corpus Christi (2024-2026)"],
  ]
  
  const wsInstrucoes = XLSX.utils.aoa_to_sheet(instrucoes)
  wsInstrucoes["!cols"] = [{ wch: 70 }]
  XLSX.utils.book_append_sheet(wb, wsInstrucoes, "Instrucoes")

  return wb
}
