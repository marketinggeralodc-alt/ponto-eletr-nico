import XLSX from "xlsx-js-style"

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

// Tipos de estilo para xlsx-js-style
interface CellStyle {
  font?: { bold?: boolean; color?: { rgb: string }; sz?: number; name?: string }
  fill?: { fgColor: { rgb: string }; patternType?: string }
  alignment?: { horizontal?: string; vertical?: string; wrapText?: boolean }
  border?: {
    top?: { style: string; color: { rgb: string } }
    bottom?: { style: string; color: { rgb: string } }
    left?: { style: string; color: { rgb: string } }
    right?: { style: string; color: { rgb: string } }
  }
  numFmt?: string
}

interface StyledCell {
  v: string | number
  t?: string
  f?: string
  s?: CellStyle
}

// Cores do tema
const CORES = {
  VERDE_ESCURO: "1B5E20",      // Cabeçalho verde escuro
  VERDE_CLARO: "C8E6C9",       // Fundo verde claro (OK)
  VERDE_TEXTO: "2E7D32",       // Texto verde
  VERMELHO_ESCURO: "B71C1C",   // Texto vermelho erro
  VERMELHO_CLARO: "FFCDD2",    // Fundo vermelho claro (erro)
  AMARELO_CLARO: "FFF9C4",     // Fundo amarelo (alerta)
  AMARELO_ESCURO: "F57F17",    // Texto amarelo escuro
  AZUL_ESCURO: "0D47A1",       // Cabeçalho azul
  AZUL_CLARO: "BBDEFB",        // Fundo azul claro
  CINZA_CLARO: "F5F5F5",       // Fundo alternado
  CINZA_MEDIO: "E0E0E0",       // Bordas
  BRANCO: "FFFFFF",
  PRETO: "000000",
  ROXO_CLARO: "E1BEE7",        // Feriado/Fim de semana
  ROXO_ESCURO: "6A1B9A",       // Texto feriado
}

// Borda padrão
const BORDA_FINA: CellStyle["border"] = {
  top: { style: "thin", color: { rgb: CORES.CINZA_MEDIO } },
  bottom: { style: "thin", color: { rgb: CORES.CINZA_MEDIO } },
  left: { style: "thin", color: { rgb: CORES.CINZA_MEDIO } },
  right: { style: "thin", color: { rgb: CORES.CINZA_MEDIO } },
}

// Opções de justificativa de falta
const JUSTIFICATIVAS = [
  "ATESTADO MÉDICO",
  "FÉRIAS",
  "FOLGA COMPENSATÓRIA",
  "LICENÇA",
  "FALTA JUSTIFICADA",
  "FALTA NÃO JUSTIFICADA",
  "TRABALHO EXTERNO",
  "HOME OFFICE",
  "OUTRO"
]

// Gera tabela de dados completa com fórmulas, estilos e cores
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
  
  // Lista de feriados do mês para referência nas fórmulas
  const feriadosDoMes: number[] = []
  for (let d = 1; d <= diasNoMes; d++) {
    if (verificarFeriado(d, mes, ano).isFeriado) {
      feriadosDoMes.push(d)
    }
  }
  
  // Agrupa dados por funcionário
  const funcionarios = [...new Set(dados.map(d => d.nome))]
  
  // Cria uma aba completa para cada funcionário
  for (const func of funcionarios) {
    const dadosFunc = dados.filter(d => d.nome === func)
    const dadosPorDia: Record<number, DadosPonto> = {}
    for (const d of dadosFunc) {
      dadosPorDia[d.dia] = d
    }
    
    // Estrutura da planilha com estilos
    const ws: XLSX.WorkSheet = {}
    
    // Linha 1: Título
    ws["A1"] = { 
      v: `FOLHA DE PONTO - ${mesAno}`, 
      s: { 
        font: { bold: true, sz: 16, color: { rgb: CORES.BRANCO }, name: "Arial" },
        fill: { fgColor: { rgb: CORES.AZUL_ESCURO }, patternType: "solid" },
        alignment: { horizontal: "center", vertical: "center" },
        border: BORDA_FINA
      } 
    }
    
    // Linha 3: Nome e Função
    ws["A3"] = { v: "NOME:", s: { font: { bold: true }, border: BORDA_FINA, fill: { fgColor: { rgb: CORES.CINZA_CLARO }, patternType: "solid" } } }
    ws["B3"] = { v: func.toUpperCase(), s: { font: { bold: true, sz: 12 }, border: BORDA_FINA, fill: { fgColor: { rgb: CORES.AMARELO_CLARO }, patternType: "solid" } } }
    ws["D3"] = { v: "FUNÇÃO:", s: { font: { bold: true }, border: BORDA_FINA, fill: { fgColor: { rgb: CORES.CINZA_CLARO }, patternType: "solid" } } }
    ws["E3"] = { v: "", s: { border: BORDA_FINA, fill: { fgColor: { rgb: CORES.AMARELO_CLARO }, patternType: "solid" } } }
    ws["G3"] = { v: "JORNADA:", s: { font: { bold: true }, border: BORDA_FINA, fill: { fgColor: { rgb: CORES.CINZA_CLARO }, patternType: "solid" } } }
    ws["H3"] = { v: "08:00", s: { border: BORDA_FINA, fill: { fgColor: { rgb: CORES.AMARELO_CLARO }, patternType: "solid" } } }
    
    // Linha 5: Cabeçalho da tabela
    const cabecalhos = ["DIA", "SEMANA", "ENTRADA 1", "SAÍDA ALM.", "ENTRADA 2", "SAÍDA", "TOTAL", "EXTRAS", "STATUS", "JUSTIFICATIVA", "ALERTA"]
    const colLetras = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K"]
    
    for (let i = 0; i < cabecalhos.length; i++) {
      ws[`${colLetras[i]}5`] = {
        v: cabecalhos[i],
        s: {
          font: { bold: true, color: { rgb: CORES.BRANCO }, sz: 10, name: "Arial" },
          fill: { fgColor: { rgb: CORES.VERDE_ESCURO }, patternType: "solid" },
          alignment: { horizontal: "center", vertical: "center", wrapText: true },
          border: BORDA_FINA
        }
      }
    }
    
    // Linha de início dos dados
    const linhaInicioDados = 6
    
    // Preenche cada dia do mês
    for (let dia = 1; dia <= diasNoMes; dia++) {
      const row = linhaInicioDados + dia - 1
      const dataAtual = new Date(ano, mes - 1, dia)
      const diaSemana = dataAtual.getDay()
      const diasSemanaTexto = ["DOM", "SEG", "TER", "QUA", "QUI", "SEX", "SAB"]
      const diaSemanaStr = diasSemanaTexto[diaSemana]
      
      const registro = dadosPorDia[dia]
      const feriado = verificarFeriado(dia, mes, ano)
      const isDomingo = diaSemana === 0
      const isSabado = diaSemana === 6
      const isFimDeSemanaOuFeriado = isDomingo || isSabado || feriado.isFeriado
      
      // Determina se é dia sem registro (possível falta)
      const semRegistro = !registro && !isFimDeSemanaOuFeriado
      
      // Define cor de fundo base
      let corFundo = dia % 2 === 0 ? CORES.CINZA_CLARO : CORES.BRANCO
      let corTexto = CORES.PRETO
      let statusTexto = ""
      let alertaTexto = ""
      
      if (isFimDeSemanaOuFeriado) {
        corFundo = CORES.ROXO_CLARO
        if (feriado.isFeriado) {
          statusTexto = `FERIADO: ${feriado.nome}`
        } else if (isDomingo) {
          statusTexto = "DOMINGO"
        } else {
          statusTexto = "SÁBADO"
        }
      } else if (semRegistro) {
        corFundo = CORES.VERMELHO_CLARO
        corTexto = CORES.VERMELHO_ESCURO
        statusTexto = "SEM REGISTRO"
      }
      
      // Horários
      const entrada1 = registro?.entrada1 || ""
      const saida1 = registro?.saida1 || ""
      const entrada2 = registro?.entrada2 || ""
      const saida2 = registro?.saida2 || ""
      
      // Verifica registros incompletos
      if (registro && registro.status.includes("Incompleto")) {
        corFundo = CORES.AMARELO_CLARO
        corTexto = CORES.AMARELO_ESCURO
        alertaTexto = registro.status
      }
      
      // Verifica almoço curto
      if (saida1 && entrada2) {
        const duracaoAlmoco = calcularDuracaoAlmoco(saida1, entrada2)
        if (duracaoAlmoco > 0 && duracaoAlmoco < 60) {
          if (alertaTexto) alertaTexto += " | "
          alertaTexto += `ALMOÇO ${minutosParaHorario(duracaoAlmoco)}`
          if (corFundo !== CORES.VERMELHO_CLARO) {
            corFundo = CORES.AMARELO_CLARO
          }
        }
      }
      
      // Se tudo OK em dia normal com registro completo
      if (registro && !registro.status.includes("Incompleto") && !isFimDeSemanaOuFeriado && !alertaTexto) {
        corFundo = CORES.VERDE_CLARO
        corTexto = CORES.VERDE_TEXTO
        statusTexto = "OK"
      }
      
      // Estilo base para a linha
      const estiloBase: CellStyle = {
        font: { color: { rgb: corTexto }, sz: 10, name: "Arial" },
        fill: { fgColor: { rgb: corFundo }, patternType: "solid" },
        alignment: { horizontal: "center", vertical: "center" },
        border: BORDA_FINA
      }
      
      // Coluna A: DIA
      ws[`A${row}`] = { v: dia, s: { ...estiloBase, font: { ...estiloBase.font, bold: true } } }
      
      // Coluna B: DIA SEMANA
      ws[`B${row}`] = { v: diaSemanaStr, s: estiloBase }
      
      // Coluna C: ENTRADA 1
      ws[`C${row}`] = { v: entrada1, s: estiloBase }
      
      // Coluna D: SAÍDA ALMOÇO
      ws[`D${row}`] = { v: saida1, s: estiloBase }
      
      // Coluna E: ENTRADA 2
      ws[`E${row}`] = { v: entrada2, s: estiloBase }
      
      // Coluna F: SAÍDA
      ws[`F${row}`] = { v: saida2, s: estiloBase }
      
      // Coluna G: TOTAL HORAS (fórmula Excel)
      // Fórmula: SE todos os horários preenchidos, calcular (F-C)-(E-D) para dias com almoço
      // Ou simplesmente (F-C) se não tiver almoço registrado
      const formulaTotal = `=SE(E(C${row}<>"";F${row}<>"");SE(E(D${row}<>"";E${row}<>"");(F${row}-C${row})-(E${row}-D${row});F${row}-C${row});"")`
      ws[`G${row}`] = { 
        f: formulaTotal, 
        s: { 
          ...estiloBase, 
          numFmt: "[h]:mm",
          font: { ...estiloBase.font, bold: true }
        } 
      }
      
      // Coluna H: HORAS EXTRAS (fórmula Excel)
      // Se fim de semana/feriado: todas as horas são extras
      // Se dia normal: horas - jornada (8h = $H$3)
      let formulaExtras: string
      if (isFimDeSemanaOuFeriado) {
        formulaExtras = `=SE(G${row}<>"";G${row};"")`
      } else {
        formulaExtras = `=SE(G${row}<>"";G${row}-$H$3;"")`
      }
      ws[`H${row}`] = { 
        f: formulaExtras, 
        s: { 
          ...estiloBase, 
          numFmt: "[h]:mm;-[h]:mm",
          font: { ...estiloBase.font, bold: true }
        } 
      }
      
      // Coluna I: STATUS
      ws[`I${row}`] = { v: statusTexto, s: estiloBase }
      
      // Coluna J: JUSTIFICATIVA (dropdown será simulado com validação - o usuário pode digitar)
      ws[`J${row}`] = { 
        v: "", 
        s: { 
          ...estiloBase, 
          fill: { fgColor: { rgb: CORES.BRANCO }, patternType: "solid" },
          font: { color: { rgb: CORES.PRETO }, sz: 9, name: "Arial" }
        } 
      }
      
      // Coluna K: ALERTA
      const estiloAlerta: CellStyle = {
        ...estiloBase,
        font: { color: { rgb: alertaTexto ? CORES.VERMELHO_ESCURO : CORES.PRETO }, sz: 9, bold: !!alertaTexto, name: "Arial" }
      }
      ws[`K${row}`] = { v: alertaTexto, s: estiloAlerta }
    }
    
    // Linha de totais
    const linhaTotais = linhaInicioDados + diasNoMes + 1
    
    // Estilo para totais
    const estiloTotais: CellStyle = {
      font: { bold: true, color: { rgb: CORES.BRANCO }, sz: 11, name: "Arial" },
      fill: { fgColor: { rgb: CORES.AZUL_ESCURO }, patternType: "solid" },
      alignment: { horizontal: "center", vertical: "center" },
      border: BORDA_FINA
    }
    
    ws[`A${linhaTotais}`] = { v: "", s: estiloTotais }
    ws[`B${linhaTotais}`] = { v: "", s: estiloTotais }
    ws[`C${linhaTotais}`] = { v: "", s: estiloTotais }
    ws[`D${linhaTotais}`] = { v: "", s: estiloTotais }
    ws[`E${linhaTotais}`] = { v: "", s: estiloTotais }
    ws[`F${linhaTotais}`] = { v: "TOTAIS:", s: estiloTotais }
    
    // Total de horas trabalhadas (fórmula SOMA)
    ws[`G${linhaTotais}`] = { 
      f: `=SOMA(G${linhaInicioDados}:G${linhaInicioDados + diasNoMes - 1})`, 
      s: { ...estiloTotais, numFmt: "[h]:mm" } 
    }
    
    // Total de horas extras (fórmula SOMA com filtro de positivos)
    ws[`H${linhaTotais}`] = { 
      f: `=SOMA(H${linhaInicioDados}:H${linhaInicioDados + diasNoMes - 1})`, 
      s: { ...estiloTotais, numFmt: "[h]:mm;-[h]:mm" } 
    }
    
    ws[`I${linhaTotais}`] = { v: "", s: estiloTotais }
    ws[`J${linhaTotais}`] = { v: "", s: estiloTotais }
    ws[`K${linhaTotais}`] = { v: "", s: estiloTotais }
    
    // Seção de Observações
    const linhaObs = linhaTotais + 2
    
    ws[`A${linhaObs}`] = { 
      v: "OBSERVAÇÕES E JUSTIFICATIVAS:", 
      s: { 
        font: { bold: true, sz: 11, name: "Arial" },
        fill: { fgColor: { rgb: CORES.CINZA_CLARO }, patternType: "solid" },
        border: BORDA_FINA
      } 
    }
    
    // Lista de justificativas possíveis
    ws[`A${linhaObs + 1}`] = { 
      v: `Opções: ${JUSTIFICATIVAS.join(", ")}`, 
      s: { 
        font: { italic: true, sz: 9, color: { rgb: "666666" }, name: "Arial" },
        alignment: { wrapText: true }
      } 
    }
    
    // Espaço para observações manuais
    for (let i = 0; i < 5; i++) {
      ws[`A${linhaObs + 3 + i}`] = { 
        v: "", 
        s: { 
          border: BORDA_FINA,
          fill: { fgColor: { rgb: CORES.BRANCO }, patternType: "solid" }
        } 
      }
    }
    
    // Assinatura
    const linhaAss = linhaObs + 10
    ws[`D${linhaAss}`] = { v: "________________________________", s: { alignment: { horizontal: "center" } } }
    ws[`D${linhaAss + 1}`] = { v: "Assinatura do Funcionário", s: { font: { sz: 9 }, alignment: { horizontal: "center" } } }
    ws[`H${linhaAss}`] = { v: "________________________________", s: { alignment: { horizontal: "center" } } }
    ws[`H${linhaAss + 1}`] = { v: "Assinatura do Responsável", s: { font: { sz: 9 }, alignment: { horizontal: "center" } } }
    
    // Legenda
    const linhaLegenda = linhaAss + 4
    ws[`A${linhaLegenda}`] = { 
      v: "LEGENDA:", 
      s: { font: { bold: true, sz: 10 } } 
    }
    
    // Cores da legenda
    const legendas = [
      { cor: CORES.VERDE_CLARO, texto: "Dia OK" },
      { cor: CORES.VERMELHO_CLARO, texto: "Sem registro / Falta" },
      { cor: CORES.AMARELO_CLARO, texto: "Registro incompleto / Alerta" },
      { cor: CORES.ROXO_CLARO, texto: "Fim de semana / Feriado" },
    ]
    
    for (let i = 0; i < legendas.length; i++) {
      const col = colLetras[i * 2]
      const colTexto = colLetras[i * 2 + 1]
      ws[`${col}${linhaLegenda + 1}`] = { 
        v: "■", 
        s: { 
          font: { color: { rgb: legendas[i].cor }, sz: 14 },
          alignment: { horizontal: "center" }
        } 
      }
      ws[`${colTexto}${linhaLegenda + 1}`] = { 
        v: legendas[i].texto, 
        s: { font: { sz: 9 } } 
      }
    }
    
    // Define range da planilha
    ws["!ref"] = `A1:K${linhaLegenda + 2}`
    
    // Define largura das colunas
    ws["!cols"] = [
      { wch: 5 },   // A: DIA
      { wch: 8 },   // B: SEMANA
      { wch: 11 },  // C: ENTRADA 1
      { wch: 11 },  // D: SAÍDA ALM
      { wch: 11 },  // E: ENTRADA 2
      { wch: 11 },  // F: SAÍDA
      { wch: 10 },  // G: TOTAL
      { wch: 10 },  // H: EXTRAS
      { wch: 22 },  // I: STATUS
      { wch: 20 },  // J: JUSTIFICATIVA
      { wch: 25 },  // K: ALERTA
    ]
    
    // Altura das linhas
    ws["!rows"] = [
      { hpt: 25 }, // Título
    ]
    
    // Merge para o título
    ws["!merges"] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 10 } }, // Título
      { s: { r: 2, c: 1 }, e: { r: 2, c: 2 } }, // Nome
      { s: { r: 2, c: 4 }, e: { r: 2, c: 5 } }, // Função
      { s: { r: linhaObs - 1, c: 0 }, e: { r: linhaObs - 1, c: 10 } }, // Observações título
      { s: { r: linhaObs, c: 0 }, e: { r: linhaObs, c: 10 } }, // Opções justificativas
    ]
    
    // Limita nome da aba a 31 caracteres
    const nomeAba = func.substring(0, 31).toUpperCase()
    XLSX.utils.book_append_sheet(wb, ws, nomeAba)
  }
  
  // Cria aba de RESUMO com fórmulas
  criarAbaResumo(wb, funcionarios, mesAno, diasNoMes)
  
  // Cria aba de dados para PROCV
  criarAbaDadosPROCV(wb, dados)
  
  // Cria aba de instruções
  criarAbaInstrucoes(wb)

  return wb
}

// Cria aba de resumo com referências às abas dos funcionários
function criarAbaResumo(wb: XLSX.WorkBook, funcionarios: string[], mesAno: string, diasNoMes: number): void {
  const ws: XLSX.WorkSheet = {}
  
  // Título
  ws["A1"] = { 
    v: `RESUMO DE HORAS EXTRAS - ${mesAno}`, 
    s: { 
      font: { bold: true, sz: 16, color: { rgb: CORES.BRANCO }, name: "Arial" },
      fill: { fgColor: { rgb: CORES.AZUL_ESCURO }, patternType: "solid" },
      alignment: { horizontal: "center", vertical: "center" },
      border: BORDA_FINA
    } 
  }
  
  // Cabeçalhos
  const cabecalhos = ["FUNCIONÁRIO", "TOTAL HORAS", "HORAS EXTRAS", "DIAS FALTANTES", "OBSERVAÇÕES"]
  const colLetras = ["A", "B", "C", "D", "E"]
  
  for (let i = 0; i < cabecalhos.length; i++) {
    ws[`${colLetras[i]}3`] = {
      v: cabecalhos[i],
      s: {
        font: { bold: true, color: { rgb: CORES.BRANCO }, sz: 11, name: "Arial" },
        fill: { fgColor: { rgb: CORES.VERDE_ESCURO }, patternType: "solid" },
        alignment: { horizontal: "center", vertical: "center" },
        border: BORDA_FINA
      }
    }
  }
  
  // Dados de cada funcionário com fórmulas referenciando as abas
  for (let i = 0; i < funcionarios.length; i++) {
    const row = 4 + i
    const nomeAba = funcionarios[i].substring(0, 31).toUpperCase()
    const linhaInicioDados = 6
    const linhaTotais = linhaInicioDados + diasNoMes + 1
    
    const corFundo = i % 2 === 0 ? CORES.CINZA_CLARO : CORES.BRANCO
    const estiloBase: CellStyle = {
      font: { sz: 10, name: "Arial" },
      fill: { fgColor: { rgb: corFundo }, patternType: "solid" },
      alignment: { horizontal: "center", vertical: "center" },
      border: BORDA_FINA
    }
    
    // Nome
    ws[`A${row}`] = { v: nomeAba, s: { ...estiloBase, font: { ...estiloBase.font, bold: true } } }
    
    // Total de horas (referência à aba do funcionário)
    ws[`B${row}`] = { 
      f: `='${nomeAba}'!G${linhaTotais}`, 
      s: { ...estiloBase, numFmt: "[h]:mm" } 
    }
    
    // Horas extras
    ws[`C${row}`] = { 
      f: `='${nomeAba}'!H${linhaTotais}`, 
      s: { ...estiloBase, numFmt: "[h]:mm;-[h]:mm" } 
    }
    
    // Dias faltantes (conta células "SEM REGISTRO")
    ws[`D${row}`] = { 
      f: `=CONT.SE('${nomeAba}'!I${linhaInicioDados}:I${linhaInicioDados + diasNoMes - 1};"SEM REGISTRO")`, 
      s: estiloBase 
    }
    
    // Observações
    ws[`E${row}`] = { v: "", s: { ...estiloBase, fill: { fgColor: { rgb: CORES.BRANCO }, patternType: "solid" } } }
  }
  
  // Linha de totais
  const rowTotal = 4 + funcionarios.length + 1
  const estiloTotais: CellStyle = {
    font: { bold: true, color: { rgb: CORES.BRANCO }, sz: 11, name: "Arial" },
    fill: { fgColor: { rgb: CORES.AZUL_ESCURO }, patternType: "solid" },
    alignment: { horizontal: "center", vertical: "center" },
    border: BORDA_FINA
  }
  
  ws[`A${rowTotal}`] = { v: "TOTAL GERAL:", s: estiloTotais }
  ws[`B${rowTotal}`] = { f: `=SOMA(B4:B${rowTotal - 2})`, s: { ...estiloTotais, numFmt: "[h]:mm" } }
  ws[`C${rowTotal}`] = { f: `=SOMA(C4:C${rowTotal - 2})`, s: { ...estiloTotais, numFmt: "[h]:mm;-[h]:mm" } }
  ws[`D${rowTotal}`] = { f: `=SOMA(D4:D${rowTotal - 2})`, s: estiloTotais }
  ws[`E${rowTotal}`] = { v: "", s: estiloTotais }
  
  // Define range e colunas
  ws["!ref"] = `A1:E${rowTotal + 1}`
  ws["!cols"] = [
    { wch: 25 },  // FUNCIONÁRIO
    { wch: 15 },  // TOTAL HORAS
    { wch: 15 },  // HORAS EXTRAS
    { wch: 15 },  // DIAS FALTANTES
    { wch: 30 },  // OBSERVAÇÕES
  ]
  
  ws["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 4 } }, // Título
  ]
  
  XLSX.utils.book_append_sheet(wb, ws, "RESUMO")
}

// Cria aba com dados para PROCV
function criarAbaDadosPROCV(wb: XLSX.WorkBook, dados: DadosPonto[]): void {
  const dadosOrdenados = [...dados].sort((a, b) => {
    const nomeCompare = a.nome.localeCompare(b.nome)
    if (nomeCompare !== 0) return nomeCompare
    return a.dia - b.dia
  })
  
  const ws: XLSX.WorkSheet = {}
  
  // Cabeçalhos
  const cabecalhos = ["FUNCIONÁRIO", "DIA", "CHAVE_PROCV", "ENTRADA_1", "SAÍDA_1", "ENTRADA_2", "SAÍDA_2", "STATUS"]
  const colLetras = ["A", "B", "C", "D", "E", "F", "G", "H"]
  
  for (let i = 0; i < cabecalhos.length; i++) {
    ws[`${colLetras[i]}1`] = {
      v: cabecalhos[i],
      s: {
        font: { bold: true, color: { rgb: CORES.BRANCO }, sz: 10, name: "Arial" },
        fill: { fgColor: { rgb: CORES.VERDE_ESCURO }, patternType: "solid" },
        alignment: { horizontal: "center" },
        border: BORDA_FINA
      }
    }
  }
  
  // Dados
  for (let i = 0; i < dadosOrdenados.length; i++) {
    const d = dadosOrdenados[i]
    const row = i + 2
    const corFundo = i % 2 === 0 ? CORES.CINZA_CLARO : CORES.BRANCO
    const estiloBase: CellStyle = {
      font: { sz: 10, name: "Arial" },
      fill: { fgColor: { rgb: corFundo }, patternType: "solid" },
      alignment: { horizontal: "center" },
      border: BORDA_FINA
    }
    
    ws[`A${row}`] = { v: d.nome.toUpperCase(), s: estiloBase }
    ws[`B${row}`] = { v: d.dia, s: estiloBase }
    ws[`C${row}`] = { v: `${d.nome.toUpperCase()}_${d.dia}`, s: estiloBase }
    ws[`D${row}`] = { v: d.entrada1, s: estiloBase }
    ws[`E${row}`] = { v: d.saida1, s: estiloBase }
    ws[`F${row}`] = { v: d.entrada2, s: estiloBase }
    ws[`G${row}`] = { v: d.saida2, s: estiloBase }
    ws[`H${row}`] = { v: d.status, s: estiloBase }
  }
  
  ws["!ref"] = `A1:H${dadosOrdenados.length + 1}`
  ws["!cols"] = [
    { wch: 20 }, { wch: 6 }, { wch: 25 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 20 }
  ]
  
  XLSX.utils.book_append_sheet(wb, ws, "Dados_PROCV")
}

// Cria aba de instruções
function criarAbaInstrucoes(wb: XLSX.WorkBook): void {
  const ws: XLSX.WorkSheet = {}
  
  const instrucoes = [
    { texto: "INSTRUÇÕES DE USO", estilo: "titulo" },
    { texto: "", estilo: "normal" },
    { texto: "ESTRUTURA DA PLANILHA:", estilo: "subtitulo" },
    { texto: "• Uma aba para cada funcionário com horários e cálculos", estilo: "normal" },
    { texto: "• Aba RESUMO com totais de todos os funcionários", estilo: "normal" },
    { texto: "• Aba Dados_PROCV para uso com fórmulas PROCV", estilo: "normal" },
    { texto: "", estilo: "normal" },
    { texto: "CORES E SIGNIFICADOS:", estilo: "subtitulo" },
    { texto: "• VERDE: Dia com registro completo e OK", estilo: "normal" },
    { texto: "• VERMELHO: Dia sem registro (possível falta)", estilo: "normal" },
    { texto: "• AMARELO: Registro incompleto ou alerta de almoço", estilo: "normal" },
    { texto: "• ROXO: Fim de semana ou feriado", estilo: "normal" },
    { texto: "", estilo: "normal" },
    { texto: "CÁLCULOS AUTOMÁTICOS:", estilo: "subtitulo" },
    { texto: "• Total de horas = (Saída - Entrada 1) - (Entrada 2 - Saída Almoço)", estilo: "normal" },
    { texto: "• Horas extras em dias normais = Total - Jornada (8h)", estilo: "normal" },
    { texto: "• Horas extras em fins de semana/feriados = Total (todas as horas)", estilo: "normal" },
    { texto: "", estilo: "normal" },
    { texto: "JUSTIFICATIVAS DISPONÍVEIS:", estilo: "subtitulo" },
    ...JUSTIFICATIVAS.map(j => ({ texto: `• ${j}`, estilo: "normal" })),
    { texto: "", estilo: "normal" },
    { texto: "FERIADOS CONSIDERADOS:", estilo: "subtitulo" },
    { texto: "• 01/01: Confraternização Universal", estilo: "normal" },
    { texto: "• 21/04: Tiradentes", estilo: "normal" },
    { texto: "• 01/05: Dia do Trabalho", estilo: "normal" },
    { texto: "• 07/09: Independência do Brasil", estilo: "normal" },
    { texto: "• 12/10: Nossa Senhora Aparecida", estilo: "normal" },
    { texto: "• 02/11: Finados", estilo: "normal" },
    { texto: "• 15/11: Proclamação da República", estilo: "normal" },
    { texto: "• 25/12: Natal", estilo: "normal" },
    { texto: "• Carnaval, Sexta-feira Santa, Corpus Christi (móveis)", estilo: "normal" },
    { texto: "", estilo: "normal" },
    { texto: "DICAS:", estilo: "subtitulo" },
    { texto: "• A jornada padrão pode ser alterada na célula H3 de cada aba", estilo: "normal" },
    { texto: "• Use a coluna JUSTIFICATIVA para explicar faltas", estilo: "normal" },
    { texto: "• As fórmulas são preservadas ao editar os horários", estilo: "normal" },
  ]
  
  for (let i = 0; i < instrucoes.length; i++) {
    const item = instrucoes[i]
    let estilo: CellStyle = { font: { sz: 10, name: "Arial" } }
    
    if (item.estilo === "titulo") {
      estilo = { 
        font: { bold: true, sz: 14, color: { rgb: CORES.BRANCO }, name: "Arial" },
        fill: { fgColor: { rgb: CORES.AZUL_ESCURO }, patternType: "solid" }
      }
    } else if (item.estilo === "subtitulo") {
      estilo = { 
        font: { bold: true, sz: 11, name: "Arial" },
        fill: { fgColor: { rgb: CORES.CINZA_CLARO }, patternType: "solid" }
      }
    }
    
    ws[`A${i + 1}`] = { v: item.texto, s: estilo }
  }
  
  ws["!ref"] = `A1:A${instrucoes.length}`
  ws["!cols"] = [{ wch: 70 }]
  
  XLSX.utils.book_append_sheet(wb, ws, "Instrucoes")
}
