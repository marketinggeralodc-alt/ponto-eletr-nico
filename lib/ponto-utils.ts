import * as XLSX from "xlsx"
import ExcelJS from "exceljs"

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
const FERIADOS_FIXOS = [
  "01/01", "21/04", "01/05", "07/09", "12/10", "02/11", "15/11", "25/12",
]

const FERIADOS_MOVEIS: Record<number, string[]> = {
  2024: ["12/02", "13/02", "29/03", "30/05"],
  2025: ["03/03", "04/03", "18/04", "19/06"],
  2026: ["16/02", "17/02", "03/04", "04/06"],
}

const FERIADOS_NOMES: Record<string, string> = {
  "01/01": "Confraternização",
  "21/04": "Tiradentes",
  "01/05": "Dia do Trabalho",
  "07/09": "Independência",
  "12/10": "N. Sra. Aparecida",
  "02/11": "Finados",
  "15/11": "Proclamação Rep.",
  "25/12": "Natal",
}

export function verificarFeriado(dia: number, mes: number, ano: number): { isFeriado: boolean; nome?: string } {
  const dataStr = `${String(dia).padStart(2, "0")}/${String(mes).padStart(2, "0")}`
  
  if (FERIADOS_FIXOS.includes(dataStr)) {
    return { isFeriado: true, nome: FERIADOS_NOMES[dataStr] || "Feriado" }
  }
  
  const feriadosMoveis = FERIADOS_MOVEIS[ano] || []
  if (feriadosMoveis.includes(dataStr)) {
    return { isFeriado: true, nome: "Feriado Móvel" }
  }
  
  return { isFeriado: false }
}

export function calcularDuracaoAlmoco(saida1: string, entrada2: string): number {
  if (!saida1 || !entrada2) return 0
  const minutosSaida = horarioParaMinutos(saida1)
  const minutosEntrada = horarioParaMinutos(entrada2)
  if (minutosSaida < 0 || minutosEntrada < 0) return 0
  return minutosEntrada - minutosSaida
}

export function calcularHorasTrabalhadas(entrada1: string, saida1: string, entrada2: string, saida2: string): number {
  let total = 0
  
  if (entrada1 && saida1) {
    const min1 = horarioParaMinutos(entrada1)
    const min2 = horarioParaMinutos(saida1)
    if (min1 >= 0 && min2 >= 0) total += min2 - min1
  }
  
  if (entrada2 && saida2) {
    const min3 = horarioParaMinutos(entrada2)
    const min4 = horarioParaMinutos(saida2)
    if (min3 >= 0 && min4 >= 0) total += min4 - min3
  }
  
  if (entrada1 && saida2 && !saida1 && !entrada2) {
    const min1 = horarioParaMinutos(entrada1)
    const min4 = horarioParaMinutos(saida2)
    if (min1 >= 0 && min4 >= 0) total = min4 - min1
  }
  
  return total
}

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

export function normalizarNome(nome: string): string {
  return nome.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9\s]/g, "").trim()
}

function extrairHorarios(val: unknown): string[] {
  if (val === null || val === undefined) return []
  const str = String(val)
  const matches = str.match(/\d{1,2}:\d{2}/g)
  return matches || []
}

function horarioParaMinutos(horario: string): number {
  const match = horario.match(/(\d{1,2}):(\d{2})/)
  if (!match) return -1
  return parseInt(match[1]) * 60 + parseInt(match[2])
}

function processarHorarios(horarios: string[]): string[] {
  if (horarios.length === 0) return []

  const horariosMinutos = horarios
    .map((h) => ({ original: h, minutos: horarioParaMinutos(h) }))
    .filter((h) => h.minutos >= 0)
    .sort((a, b) => a.minutos - b.minutos)

  if (horariosMinutos.length === 0) return []

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

function normalizarHorario(horario: string): string {
  const match = horario.match(/(\d{1,2}):(\d{2})/)
  if (!match) return horario
  return `${match[1].padStart(2, "0")}:${match[2]}`
}

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
  let entrada1 = "", saida1 = "", entrada2 = "", saida2 = "", status = ""
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

function tentarParseHorizontal(
  linhas: unknown[][],
  debug: string[]
): { dados: DadosPonto[]; ano: number; mes: number } | null {
  debug.push("Tentando parse horizontal (dias nas colunas)")

  let anoPonto = new Date().getFullYear()
  let mesPonto = new Date().getMonth() + 1

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
          if (!diasColunas[dia]) diasColunas[dia] = j
        }
      }
      break
    }
  }

  if (linhaDias === -1) {
    debug.push("Nao encontrou linha de dias")
    return null
  }

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

      if (!nomeFuncionario) {
        const match = linhaTexto.match(/nome\s*:?\s*([a-záéíóúâêîôûãõç\s]+)/i)
        if (match && match[1]) {
          const nome = match[1].trim()
          if (nome.length > 1 && !nome.includes("dept")) nomeFuncionario = nome
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

  const dados: DadosPonto[] = []

  for (const bloco of blocosFuncionarios) {
    const horariosPorDia: Record<number, string[]> = {}

    for (let row = bloco.linhaInicio; row <= Math.min(bloco.linhaFim, bloco.linhaInicio + 10); row++) {
      const linha = linhas[row]
      if (!linha) continue

      const linhaTexto = linha.map((x) => String(x || "").toLowerCase()).join(" ")
      if (linhaTexto.includes("id") && linhaTexto.includes("nome")) break

      for (const [diaStr, colIdx] of Object.entries(diasColunas)) {
        const dia = parseInt(diaStr)
        if (colIdx >= linha.length) continue

        const horarios = extrairHorarios(linha[colIdx])
        if (horarios.length > 0) {
          if (!horariosPorDia[dia]) horariosPorDia[dia] = []
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

// ============================================================================
// GERAÇÃO DA PLANILHA COM EXCELJS
// ============================================================================

const JUSTIFICATIVAS = [
  "ATESTADO MÉDICO",
  "FÉRIAS", 
  "FOLGA",
  "LICENÇA",
  "FALTA JUSTIFICADA",
  "FALTA NÃO JUST.",
  "TRABALHO EXT.",
  "HOME OFFICE",
]

// Cores do tema
const CORES = {
  AZUL_ESCURO: "0D47A1",
  VERDE_ESCURO: "1B5E20",
  VERDE_CLARO: "C8E6C9",
  VERDE_TEXTO: "1B5E20",
  VERMELHO_CLARO: "FFCDD2",
  VERMELHO_TEXTO: "B71C1C",
  AMARELO_CLARO: "FFF9C4",
  AMARELO_TEXTO: "F57F17",
  ROXO_CLARO: "E1BEE7",
  ROXO_TEXTO: "6A1B9A",
  CINZA_CLARO: "F5F5F5",
  CINZA_BORDA: "BDBDBD",
  BRANCO: "FFFFFF",
  PRETO: "000000",
}

// Gera a planilha usando ExcelJS
export async function gerarTabelaDadosExcel(
  dados: DadosPonto[],
  mesAno: string
): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook()
  workbook.creator = "Sistema de Ponto"
  workbook.created = new Date()
  
  const [mesStr, anoStr] = mesAno.split("/")
  const mes = parseInt(mesStr) || new Date().getMonth() + 1
  const ano = parseInt(anoStr) || new Date().getFullYear()
  const diasNoMes = new Date(ano, mes, 0).getDate()
  
  const funcionarios = [...new Set(dados.map(d => d.nome))]
  const JORNADA_MINUTOS = 8 * 60
  
  // Dados para o resumo
  const resumoFuncionarios: { nome: string; totalHoras: number; totalExtras: number; diasFalta: number }[] = []
  
  // Cria uma aba para cada funcionário
  for (const func of funcionarios) {
    const dadosFunc = dados.filter(d => d.nome === func)
    const dadosPorDia: Record<number, DadosPonto> = {}
    for (const d of dadosFunc) {
      dadosPorDia[d.dia] = d
    }
    
    const nomeAba = func.substring(0, 31).toUpperCase()
    const ws = workbook.addWorksheet(nomeAba)
    
    // Configuração das colunas
    ws.columns = [
      { key: "dia", width: 6 },
      { key: "semana", width: 8 },
      { key: "entrada1", width: 11 },
      { key: "saidaAlm", width: 11 },
      { key: "entrada2", width: 11 },
      { key: "saida", width: 10 },
      { key: "total", width: 10 },
      { key: "extras", width: 10 },
      { key: "status", width: 15 },
      { key: "justificativa", width: 16 },
      { key: "observacao", width: 25 },
    ]
    
    // Linha 1: Título
    ws.mergeCells("A1:K1")
    const tituloCell = ws.getCell("A1")
    tituloCell.value = `FOLHA DE PONTO - ${mesAno}`
    tituloCell.font = { bold: true, size: 14, color: { argb: "FFFFFFFF" } }
    tituloCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + CORES.AZUL_ESCURO } }
    tituloCell.alignment = { horizontal: "center", vertical: "middle" }
    tituloCell.border = getBorder()
    ws.getRow(1).height = 25
    
    // Linha 3: Informações do funcionário
    ws.getCell("A3").value = "NOME:"
    ws.getCell("A3").font = { bold: true }
    ws.getCell("A3").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + CORES.CINZA_CLARO } }
    ws.getCell("A3").border = getBorder()
    
    ws.mergeCells("B3:C3")
    ws.getCell("B3").value = func.toUpperCase()
    ws.getCell("B3").font = { bold: true, size: 12 }
    ws.getCell("B3").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + CORES.AMARELO_CLARO } }
    ws.getCell("B3").border = getBorder()
    
    ws.getCell("D3").value = "FUNÇÃO:"
    ws.getCell("D3").font = { bold: true }
    ws.getCell("D3").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + CORES.CINZA_CLARO } }
    ws.getCell("D3").border = getBorder()
    
    ws.mergeCells("E3:F3")
    ws.getCell("E3").value = ""
    ws.getCell("E3").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + CORES.AMARELO_CLARO } }
    ws.getCell("E3").border = getBorder()
    
    ws.getCell("G3").value = "JORNADA:"
    ws.getCell("G3").font = { bold: true }
    ws.getCell("G3").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + CORES.CINZA_CLARO } }
    ws.getCell("G3").border = getBorder()
    
    // Jornada como valor de hora (8 horas = 8/24)
    ws.getCell("H3").value = 8/24  // 8 horas em formato de fração de dia do Excel
    ws.getCell("H3").numFmt = "HH:MM"
    ws.getCell("H3").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + CORES.AMARELO_CLARO } }
    ws.getCell("H3").border = getBorder()
    ws.getCell("H3").alignment = { horizontal: "center" }
    
    // Linha 5: Cabeçalhos
    const cabecalhos = ["DIA", "SEMANA", "ENTRADA 1", "SAÍDA ALM.", "ENTRADA 2", "SAÍDA", "TOTAL", "EXTRAS", "STATUS", "JUSTIFICATIVA", "OBSERVAÇÃO"]
    const headerRow = ws.getRow(5)
    cabecalhos.forEach((cab, idx) => {
      const cell = headerRow.getCell(idx + 1)
      cell.value = cab
      cell.font = { bold: true, size: 10, color: { argb: "FFFFFFFF" } }
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + CORES.VERDE_ESCURO } }
      cell.alignment = { horizontal: "center", vertical: "middle" }
      cell.border = getBorder()
    })
    headerRow.height = 20
    
    let totalMinutosMes = 0
    let totalExtrasMes = 0
    let diasFalta = 0
    
    // Dados de cada dia
    for (let dia = 1; dia <= diasNoMes; dia++) {
      const rowNum = 5 + dia
      const row = ws.getRow(rowNum)
      
      const dataAtual = new Date(ano, mes - 1, dia)
      const diaSemana = dataAtual.getDay()
      const diasSemanaTexto = ["DOM", "SEG", "TER", "QUA", "QUI", "SEX", "SAB"]
      const diaSemanaStr = diasSemanaTexto[diaSemana]
      
      const registro = dadosPorDia[dia]
      const feriado = verificarFeriado(dia, mes, ano)
      const isDomingo = diaSemana === 0
      const isSabado = diaSemana === 6
      const isFimDeSemanaOuFeriado = isDomingo || isSabado || feriado.isFeriado
      
      let entrada1 = ""
      let saida1 = ""
      let entrada2 = ""
      let saida2 = ""
      let status = ""
      let observacao = ""
      
      // Determina cor e status
      let corFundo = dia % 2 === 0 ? CORES.CINZA_CLARO : CORES.BRANCO
      let corTexto = CORES.PRETO
      
      if (registro) {
        entrada1 = registro.entrada1
        saida1 = registro.saida1
        entrada2 = registro.entrada2
        saida2 = registro.saida2
        
        // Calcula para estatísticas (ainda usamos para resumo)
        const minutosTrabalhados = calcularHorasTrabalhadas(entrada1, saida1, entrada2, saida2)
        if (minutosTrabalhados > 0) {
          totalMinutosMes += minutosTrabalhados
          if (isFimDeSemanaOuFeriado) {
            totalExtrasMes += minutosTrabalhados
          } else {
            const extras = minutosTrabalhados - JORNADA_MINUTOS
            if (extras > 0) totalExtrasMes += extras
          }
        }
        
        // Verifica almoço curto
        const duracaoAlmoco = calcularDuracaoAlmoco(saida1, entrada2)
        if (duracaoAlmoco > 0 && duracaoAlmoco < 60) {
          observacao = `ALMOÇO ${minutosParaHorario(duracaoAlmoco)}`
          corFundo = CORES.AMARELO_CLARO
          corTexto = CORES.AMARELO_TEXTO
        }
        
        // Verifica status
        if (registro.status.includes("Incompleto")) {
          status = "INCOMPLETO"
          observacao = observacao ? observacao + " | " + registro.status : registro.status
          corFundo = CORES.AMARELO_CLARO
          corTexto = CORES.AMARELO_TEXTO
        } else if (isFimDeSemanaOuFeriado) {
          status = feriado.isFeriado ? "FERIADO" : diaSemanaStr
          corFundo = CORES.ROXO_CLARO
          corTexto = CORES.ROXO_TEXTO
        } else {
          status = "OK"
          corFundo = CORES.VERDE_CLARO
          corTexto = CORES.VERDE_TEXTO
        }
      } else {
        // Sem registro
        if (feriado.isFeriado) {
          status = "FERIADO"
          corFundo = CORES.ROXO_CLARO
          corTexto = CORES.ROXO_TEXTO
        } else if (isDomingo) {
          status = "DOMINGO"
          corFundo = CORES.ROXO_CLARO
          corTexto = CORES.ROXO_TEXTO
        } else if (isSabado) {
          status = "SÁBADO"
          corFundo = CORES.ROXO_CLARO
          corTexto = CORES.ROXO_TEXTO
        } else {
          status = "SEM REGISTRO"
          corFundo = CORES.VERMELHO_CLARO
          corTexto = CORES.VERMELHO_TEXTO
          diasFalta++
        }
      }
      
      // Estilo base para as células
      const estiloBase = {
        font: { size: 10, color: { argb: "FF" + corTexto } },
        fill: { type: "pattern" as const, pattern: "solid" as const, fgColor: { argb: "FF" + corFundo } },
        alignment: { horizontal: "center" as const, vertical: "middle" as const },
        border: getBorder()
      }
      
      // Coluna A: DIA
      const cellDia = row.getCell(1)
      cellDia.value = dia
      Object.assign(cellDia, estiloBase)
      
      // Coluna B: SEMANA
      const cellSemana = row.getCell(2)
      cellSemana.value = diaSemanaStr
      Object.assign(cellSemana, estiloBase)
      
      // Coluna C: ENTRADA 1 (formato hora)
      const cellEntrada1 = row.getCell(3)
      if (entrada1) {
        const [h, m] = entrada1.split(":").map(Number)
        cellEntrada1.value = new Date(1899, 11, 30, h, m)
        cellEntrada1.numFmt = "HH:MM"
      } else {
        cellEntrada1.value = ""
      }
      Object.assign(cellEntrada1, estiloBase)
      
      // Coluna D: SAÍDA ALM (formato hora)
      const cellSaida1 = row.getCell(4)
      if (saida1) {
        const [h, m] = saida1.split(":").map(Number)
        cellSaida1.value = new Date(1899, 11, 30, h, m)
        cellSaida1.numFmt = "HH:MM"
      } else {
        cellSaida1.value = ""
      }
      Object.assign(cellSaida1, estiloBase)
      
      // Coluna E: ENTRADA 2 (formato hora)
      const cellEntrada2 = row.getCell(5)
      if (entrada2) {
        const [h, m] = entrada2.split(":").map(Number)
        cellEntrada2.value = new Date(1899, 11, 30, h, m)
        cellEntrada2.numFmt = "HH:MM"
      } else {
        cellEntrada2.value = ""
      }
      Object.assign(cellEntrada2, estiloBase)
      
      // Coluna F: SAÍDA (formato hora)
      const cellSaida2 = row.getCell(6)
      if (saida2) {
        const [h, m] = saida2.split(":").map(Number)
        cellSaida2.value = new Date(1899, 11, 30, h, m)
        cellSaida2.numFmt = "HH:MM"
      } else {
        cellSaida2.value = ""
      }
      Object.assign(cellSaida2, estiloBase)
      
      // Coluna G: TOTAL HORAS (FÓRMULA EXCEL)
      // Fórmula: SE tem entrada e saída, calcula (F-C)-(E-D) ou apenas (F-C) se não tem almoço
      const cellTotal = row.getCell(7)
      // Usa sintaxe internacional do Excel
      cellTotal.value = { 
        formula: `IF(AND(C${rowNum}<>"",F${rowNum}<>""),IF(AND(D${rowNum}<>"",E${rowNum}<>""),(F${rowNum}-C${rowNum})-(E${rowNum}-D${rowNum}),F${rowNum}-C${rowNum}),"")`,
        date1904: false
      }
      cellTotal.numFmt = "[H]:MM"
      Object.assign(cellTotal, estiloBase)
      cellTotal.font = { size: 10, bold: true, color: { argb: "FF" + corTexto } }
      
      // Coluna H: HORAS EXTRAS (FÓRMULA EXCEL)
      // Para fds/feriado: igual ao total
      // Para dia normal: total - jornada (H3 agora é valor numérico)
      const cellExtras = row.getCell(8)
      if (isFimDeSemanaOuFeriado) {
        // Fim de semana/feriado: todas as horas são extras
        cellExtras.value = { 
          formula: `IF(G${rowNum}<>"",G${rowNum},"")`,
          date1904: false
        }
      } else {
        // Dia normal: total - jornada (H3 é valor numérico de hora)
        cellExtras.value = { 
          formula: `IF(G${rowNum}<>"",G${rowNum}-$H$3,"")`,
          date1904: false
        }
      }
      cellExtras.numFmt = "[H]:MM;-[H]:MM"
      Object.assign(cellExtras, estiloBase)
      cellExtras.font = { size: 10, bold: true, color: { argb: "FF" + corTexto } }
      
      // Coluna I: STATUS
      const cellStatus = row.getCell(9)
      cellStatus.value = status
      Object.assign(cellStatus, estiloBase)
      
      // Coluna J: JUSTIFICATIVA (editável)
      const cellJust = row.getCell(10)
      cellJust.value = ""
      cellJust.font = { size: 9, color: { argb: "FF" + CORES.PRETO } }
      cellJust.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + CORES.BRANCO } }
      cellJust.alignment = { horizontal: "center", vertical: "middle" }
      cellJust.border = getBorder()
      
      // Coluna K: OBSERVAÇÃO
      const cellObs = row.getCell(11)
      cellObs.value = observacao
      Object.assign(cellObs, estiloBase)
    }
    
    // Linha de totais com FÓRMULAS
    const linhaTotais = 6 + diasNoMes
    ws.getRow(linhaTotais).height = 5 // Linha vazia
    
    const linhaTotaisReal = linhaTotais + 1
    const totaisRow = ws.getRow(linhaTotaisReal)
    const linhaInicioDados = 6
    const linhaFimDados = 5 + diasNoMes
    
    const estiloTotais = {
      font: { bold: true, size: 10, color: { argb: "FFFFFFFF" } },
      fill: { type: "pattern" as const, pattern: "solid" as const, fgColor: { argb: "FF" + CORES.AZUL_ESCURO } },
      alignment: { horizontal: "center" as const, vertical: "middle" as const },
      border: getBorder()
    }
    
    // Células vazias A-E
    for (let i = 1; i <= 5; i++) {
      const cell = totaisRow.getCell(i)
      cell.value = ""
      Object.assign(cell, estiloTotais)
    }
    
    // F: Label TOTAIS
    const cellLabelTotais = totaisRow.getCell(6)
    cellLabelTotais.value = "TOTAIS:"
    Object.assign(cellLabelTotais, estiloTotais)
    
    // G: SOMA do total de horas (FÓRMULA)
    const cellSomaTotal = totaisRow.getCell(7)
    cellSomaTotal.value = { formula: `SUM(G${linhaInicioDados}:G${linhaFimDados})` }
    cellSomaTotal.numFmt = "[H]:MM"
    Object.assign(cellSomaTotal, estiloTotais)
    
    // H: SOMA das horas extras (FÓRMULA)
    const cellSomaExtras = totaisRow.getCell(8)
    cellSomaExtras.value = { formula: `SUM(H${linhaInicioDados}:H${linhaFimDados})` }
    cellSomaExtras.numFmt = "[H]:MM;-[H]:MM"
    Object.assign(cellSomaExtras, estiloTotais)
    
    // I: Contagem de faltas (FÓRMULA COUNTIF)
    const cellFaltas = totaisRow.getCell(9)
    cellFaltas.value = { formula: `COUNTIF(I${linhaInicioDados}:I${linhaFimDados},"SEM REGISTRO")&" faltas"` }
    Object.assign(cellFaltas, estiloTotais)
    
    // J e K: vazias
    totaisRow.getCell(10).value = ""
    Object.assign(totaisRow.getCell(10), estiloTotais)
    totaisRow.getCell(11).value = ""
    Object.assign(totaisRow.getCell(11), estiloTotais)
    
    // Seção de observações
    const linhaObs = linhaTotaisReal + 2
    ws.mergeCells(`A${linhaObs}:K${linhaObs}`)
    ws.getCell(`A${linhaObs}`).value = "OBSERVAÇÕES GERAIS:"
    ws.getCell(`A${linhaObs}`).font = { bold: true, size: 11 }
    ws.getCell(`A${linhaObs}`).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + CORES.CINZA_CLARO } }
    
    ws.mergeCells(`A${linhaObs + 1}:K${linhaObs + 1}`)
    ws.getCell(`A${linhaObs + 1}`).value = `JUSTIFICATIVAS: ${JUSTIFICATIVAS.join(" | ")}`
    ws.getCell(`A${linhaObs + 1}`).font = { size: 9, italic: true }
    
    // Espaço para observações
    for (let i = 2; i <= 4; i++) {
      ws.mergeCells(`A${linhaObs + i}:K${linhaObs + i}`)
      ws.getCell(`A${linhaObs + i}`).border = getBorder()
    }
    
    // Assinaturas
    const linhaAss = linhaObs + 7
    ws.getCell(`D${linhaAss}`).value = "________________________"
    ws.getCell(`D${linhaAss}`).alignment = { horizontal: "center" }
    ws.getCell(`D${linhaAss + 1}`).value = "Assin. Funcionário"
    ws.getCell(`D${linhaAss + 1}`).font = { size: 9 }
    ws.getCell(`D${linhaAss + 1}`).alignment = { horizontal: "center" }
    
    ws.getCell(`H${linhaAss}`).value = "________________________"
    ws.getCell(`H${linhaAss}`).alignment = { horizontal: "center" }
    ws.getCell(`H${linhaAss + 1}`).value = "Assin. Responsável"
    ws.getCell(`H${linhaAss + 1}`).font = { size: 9 }
    ws.getCell(`H${linhaAss + 1}`).alignment = { horizontal: "center" }
    
    // Legenda
    const linhaLeg = linhaAss + 4
    ws.mergeCells(`A${linhaLeg}:K${linhaLeg}`)
    ws.getCell(`A${linhaLeg}`).value = "LEGENDA: Verde = OK | Vermelho = Sem Registro | Amarelo = Incompleto/Alerta | Roxo = Feriado/Fim de Semana"
    ws.getCell(`A${linhaLeg}`).font = { size: 9, italic: true }
    
    // Adiciona FORMATAÇÃO CONDICIONAL para linhas de dados
    // Isso faz com que as cores mudem automaticamente quando o usuário editar
    const linhaInicio = 6
    const linhaFim = 5 + diasNoMes
    
    // Regra 1: Se tem todas as 4 entradas preenchidas = VERDE (OK)
    ws.addConditionalFormatting({
      ref: `A${linhaInicio}:K${linhaFim}`,
      rules: [
        {
          type: "expression",
          formulae: [`AND($C${linhaInicio}<>"",$D${linhaInicio}<>"",$E${linhaInicio}<>"",$F${linhaInicio}<>"")`],
          style: {
            fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFC8E6C9" } },
            font: { color: { argb: "FF1B5E20" } }
          },
          priority: 1
        }
      ]
    })
    
    // Regra 2: Se tem entrada mas falta algum horário = AMARELO (Incompleto)
    ws.addConditionalFormatting({
      ref: `A${linhaInicio}:K${linhaFim}`,
      rules: [
        {
          type: "expression",
          formulae: [`AND(OR($C${linhaInicio}<>"",$F${linhaInicio}<>""),NOT(AND($C${linhaInicio}<>"",$D${linhaInicio}<>"",$E${linhaInicio}<>"",$F${linhaInicio}<>"")))`],
          style: {
            fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFFFF9C4" } },
            font: { color: { argb: "FFF57F17" } }
          },
          priority: 2
        }
      ]
    })
    
    // Regra 3: Se não tem nenhum horário E não é fim de semana = VERMELHO (Falta)
    ws.addConditionalFormatting({
      ref: `A${linhaInicio}:K${linhaFim}`,
      rules: [
        {
          type: "expression",
          formulae: [`AND($C${linhaInicio}="",$D${linhaInicio}="",$E${linhaInicio}="",$F${linhaInicio}="",$I${linhaInicio}<>"FERIADO",$I${linhaInicio}<>"DOMINGO",$I${linhaInicio}<>"SÁBADO",$I${linhaInicio}<>"DOM",$I${linhaInicio}<>"SAB")`],
          style: {
            fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFFFCDD2" } },
            font: { color: { argb: "FFB71C1C" } }
          },
          priority: 3
        }
      ]
    })
    
    // Guarda dados para resumo
    resumoFuncionarios.push({
      nome: func,
      totalHoras: totalMinutosMes,
      totalExtras: totalExtrasMes,
      diasFalta
    })
  }
  
  // Cria aba de RESUMO
  criarAbaResumoExcel(workbook, resumoFuncionarios, mesAno)
  
  // Cria aba de dados para PROCV
  criarAbaDadosPROCVExcel(workbook, dados)
  
  // Cria aba de instruções
  criarAbaInstrucoesExcel(workbook)
  
  // Gera o buffer
  const buffer = await workbook.xlsx.writeBuffer()
  return buffer as Buffer
}

function getBorder(): Partial<ExcelJS.Borders> {
  return {
    top: { style: "thin", color: { argb: "FF" + CORES.CINZA_BORDA } },
    left: { style: "thin", color: { argb: "FF" + CORES.CINZA_BORDA } },
    bottom: { style: "thin", color: { argb: "FF" + CORES.CINZA_BORDA } },
    right: { style: "thin", color: { argb: "FF" + CORES.CINZA_BORDA } },
  }
}

function criarAbaResumoExcel(
  workbook: ExcelJS.Workbook,
  resumoFuncionarios: { nome: string; totalHoras: number; totalExtras: number; diasFalta: number }[],
  mesAno: string
): void {
  const ws = workbook.addWorksheet("RESUMO")
  
  ws.columns = [
    { key: "funcionario", width: 25 },
    { key: "totalHoras", width: 15 },
    { key: "horasExtras", width: 15 },
    { key: "diasFalta", width: 12 },
    { key: "observacoes", width: 30 },
  ]
  
  // Título
  ws.mergeCells("A1:E1")
  const tituloCell = ws.getCell("A1")
  tituloCell.value = `RESUMO GERAL - ${mesAno}`
  tituloCell.font = { bold: true, size: 14, color: { argb: "FFFFFFFF" } }
  tituloCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + CORES.AZUL_ESCURO } }
  tituloCell.alignment = { horizontal: "center", vertical: "middle" }
  tituloCell.border = getBorder()
  ws.getRow(1).height = 25
  
  // Cabeçalhos
  const cabecalhos = ["FUNCIONÁRIO", "TOTAL HORAS", "HORAS EXTRAS", "DIAS FALTA", "OBSERVAÇÕES"]
  const headerRow = ws.getRow(3)
  cabecalhos.forEach((cab, idx) => {
    const cell = headerRow.getCell(idx + 1)
    cell.value = cab
    cell.font = { bold: true, size: 10, color: { argb: "FFFFFFFF" } }
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + CORES.VERDE_ESCURO } }
    cell.alignment = { horizontal: "center", vertical: "middle" }
    cell.border = getBorder()
  })
  
  let totalGeralHoras = 0
  let totalGeralExtras = 0
  let totalGeralFaltas = 0
  
  // Dados dos funcionários
  resumoFuncionarios.forEach((func, idx) => {
    const row = ws.getRow(4 + idx)
    const corFundo = idx % 2 === 0 ? CORES.CINZA_CLARO : CORES.BRANCO
    
    const valores = [
      func.nome.toUpperCase(),
      minutosParaHorario(func.totalHoras),
      minutosParaHorario(func.totalExtras),
      func.diasFalta,
      ""
    ]
    
    valores.forEach((val, colIdx) => {
      const cell = row.getCell(colIdx + 1)
      cell.value = val
      cell.font = { size: 10 }
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + corFundo } }
      cell.alignment = { horizontal: "center", vertical: "middle" }
      cell.border = getBorder()
    })
    
    totalGeralHoras += func.totalHoras
    totalGeralExtras += func.totalExtras
    totalGeralFaltas += func.diasFalta
  })
  
  // Linha de totais
  const linhaTotais = 4 + resumoFuncionarios.length + 1
  const totaisRow = ws.getRow(linhaTotais)
  
  const valoresTotais = [
    "TOTAL GERAL",
    minutosParaHorario(totalGeralHoras),
    minutosParaHorario(totalGeralExtras),
    totalGeralFaltas,
    ""
  ]
  
  valoresTotais.forEach((val, idx) => {
    const cell = totaisRow.getCell(idx + 1)
    cell.value = val
    cell.font = { bold: true, size: 10, color: { argb: "FFFFFFFF" } }
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + CORES.AZUL_ESCURO } }
    cell.alignment = { horizontal: "center", vertical: "middle" }
    cell.border = getBorder()
  })
}

function criarAbaDadosPROCVExcel(workbook: ExcelJS.Workbook, dados: DadosPonto[]): void {
  const ws = workbook.addWorksheet("Dados_PROCV")
  
  ws.columns = [
    { key: "funcionario", width: 20 },
    { key: "dia", width: 6 },
    { key: "chave", width: 25 },
    { key: "entrada1", width: 10 },
    { key: "saida1", width: 10 },
    { key: "entrada2", width: 10 },
    { key: "saida2", width: 10 },
    { key: "status", width: 20 },
  ]
  
  // Cabeçalhos
  const cabecalhos = ["FUNCIONÁRIO", "DIA", "CHAVE_PROCV", "ENTRADA_1", "SAÍDA_1", "ENTRADA_2", "SAÍDA_2", "STATUS"]
  const headerRow = ws.getRow(1)
  cabecalhos.forEach((cab, idx) => {
    const cell = headerRow.getCell(idx + 1)
    cell.value = cab
    cell.font = { bold: true, size: 10, color: { argb: "FFFFFFFF" } }
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + CORES.VERDE_ESCURO } }
    cell.alignment = { horizontal: "center" }
    cell.border = getBorder()
  })
  
  // Ordena dados
  const dadosOrdenados = [...dados].sort((a, b) => {
    const nomeCompare = a.nome.localeCompare(b.nome)
    if (nomeCompare !== 0) return nomeCompare
    return a.dia - b.dia
  })
  
  // Dados
  dadosOrdenados.forEach((d, idx) => {
    const row = ws.getRow(2 + idx)
    const corFundo = idx % 2 === 0 ? CORES.CINZA_CLARO : CORES.BRANCO
    
    const valores = [
      d.nome.toUpperCase(),
      d.dia,
      `${d.nome.toUpperCase()}_${d.dia}`,
      d.entrada1,
      d.saida1,
      d.entrada2,
      d.saida2,
      d.status
    ]
    
    valores.forEach((val, colIdx) => {
      const cell = row.getCell(colIdx + 1)
      cell.value = val
      cell.font = { size: 10 }
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + corFundo } }
      cell.alignment = { horizontal: "center" }
      cell.border = getBorder()
    })
  })
}

function criarAbaInstrucoesExcel(workbook: ExcelJS.Workbook): void {
  const ws = workbook.addWorksheet("Instrucoes")
  
  ws.columns = [{ key: "texto", width: 80 }]
  
  const linhas = [
    { texto: "INSTRUÇÕES DE USO", estilo: "titulo" },
    { texto: "", estilo: "normal" },
    { texto: "ESTRUTURA DA PLANILHA:", estilo: "subtitulo" },
    { texto: "- Uma aba para cada funcionário com todos os horários do mês", estilo: "normal" },
    { texto: "- Aba RESUMO com totais de horas de todos os funcionários", estilo: "normal" },
    { texto: "- Aba Dados_PROCV para integração com outras planilhas", estilo: "normal" },
    { texto: "", estilo: "normal" },
    { texto: "FÓRMULAS AUTOMÁTICAS:", estilo: "subtitulo" },
    { texto: "- As colunas TOTAL e EXTRAS contêm fórmulas que calculam automaticamente", estilo: "normal" },
    { texto: "- Ao editar os horários manualmente, os cálculos serão atualizados", estilo: "normal" },
    { texto: "- A jornada padrão está na célula H3 e pode ser alterada", estilo: "normal" },
    { texto: "- Os totais no final usam SOMA e CONT.SE para calcular automaticamente", estilo: "normal" },
    { texto: "", estilo: "normal" },
    { texto: "FORMATAÇÃO CONDICIONAL (cores automáticas):", estilo: "subtitulo" },
    { texto: "- VERDE: Quando preencher todos os 4 horários (entrada, saída almoço, entrada, saída)", estilo: "normal" },
    { texto: "- AMARELO: Quando preencher parcialmente (falta algum horário)", estilo: "normal" },
    { texto: "- VERMELHO: Quando dia útil estiver sem nenhum horário", estilo: "normal" },
    { texto: "- ROXO: Feriados e fins de semana (não muda com edição)", estilo: "normal" },
    { texto: "", estilo: "normal" },
    { texto: "CÁLCULOS (fórmulas Excel):", estilo: "subtitulo" },
    { texto: "- Total de horas = (Saída - Entrada1) - (Entrada2 - Saída Almoço)", estilo: "normal" },
    { texto: "- Horas extras em dias normais = Total - Jornada (célula H3)", estilo: "normal" },
    { texto: "- Horas extras em fins de semana/feriados = Total (todas são extras)", estilo: "normal" },
    { texto: "- Valores NEGATIVOS indicam horas faltantes (devendo)", estilo: "normal" },
    { texto: "", estilo: "normal" },
    { texto: "COMO PREENCHER MANUALMENTE:", estilo: "subtitulo" },
    { texto: "- Digite os horários no formato HH:MM (ex: 08:00, 12:00, 13:00, 17:00)", estilo: "normal" },
    { texto: "- O Excel converterá automaticamente para o formato de hora", estilo: "normal" },
    { texto: "- Os cálculos de TOTAL e EXTRAS serão atualizados automaticamente", estilo: "normal" },
    { texto: "- As CORES mudarão automaticamente conforme você preencher", estilo: "normal" },
    { texto: "", estilo: "normal" },
    { texto: "COMO ALTERAR A JORNADA:", estilo: "subtitulo" },
    { texto: "- A jornada está na célula H3 de cada aba de funcionário", estilo: "normal" },
    { texto: "- Para 6 horas: digite 06:00 ou 0,25 (6/24)", estilo: "normal" },
    { texto: "- Para 8 horas: digite 08:00 ou 0,333 (8/24)", estilo: "normal" },
    { texto: "- Ao alterar, todas as horas extras serão recalculadas", estilo: "normal" },
    { texto: "", estilo: "normal" },
    { texto: "JUSTIFICATIVAS DISPONÍVEIS:", estilo: "subtitulo" },
    ...JUSTIFICATIVAS.map(j => ({ texto: `- ${j}`, estilo: "normal" as const })),
    { texto: "", estilo: "normal" },
    { texto: "FERIADOS CONSIDERADOS:", estilo: "subtitulo" },
    { texto: "- 01/01: Confraternização Universal", estilo: "normal" },
    { texto: "- 21/04: Tiradentes", estilo: "normal" },
    { texto: "- 01/05: Dia do Trabalho", estilo: "normal" },
    { texto: "- 07/09: Independência do Brasil", estilo: "normal" },
    { texto: "- 12/10: Nossa Senhora Aparecida", estilo: "normal" },
    { texto: "- 02/11: Finados", estilo: "normal" },
    { texto: "- 15/11: Proclamação da República", estilo: "normal" },
    { texto: "- 25/12: Natal", estilo: "normal" },
    { texto: "- Carnaval, Sexta-feira Santa, Corpus Christi (datas móveis)", estilo: "normal" },
  ]
  
  linhas.forEach((item, idx) => {
    const row = ws.getRow(idx + 1)
    const cell = row.getCell(1)
    cell.value = item.texto
    
    if (item.estilo === "titulo") {
      cell.font = { bold: true, size: 14, color: { argb: "FFFFFFFF" } }
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + CORES.AZUL_ESCURO } }
    } else if (item.estilo === "subtitulo") {
      cell.font = { bold: true, size: 11 }
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + CORES.CINZA_CLARO } }
    } else {
      cell.font = { size: 10 }
    }
  })
}

// Função wrapper para manter compatibilidade (retorna XLSX.WorkBook para leitura)
export function gerarTabelaDados(
  dados: DadosPonto[],
  mesAno: string
): XLSX.WorkBook {
  // Retorna um workbook vazio - a função real é gerarTabelaDadosExcel
  // Esta função é mantida apenas para compatibilidade de tipos
  return XLSX.utils.book_new()
}
