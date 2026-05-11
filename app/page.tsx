"use client"

import { useState, useCallback } from "react"
import * as XLSX from "xlsx"
import { FileUpload } from "@/components/file-upload"
import { Button } from "@/components/ui/button"
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card"
import { processarRelogio, gerarTabelaDados, type DadosPonto } from "@/lib/ponto-utils"
import { Download, Clock, FileSpreadsheet, AlertCircle, CheckCircle2, Loader2, Users, Calendar, Info, ChevronDown, ChevronUp } from "lucide-react"

type Step = "upload" | "preview" | "result"

interface ProcessingResult {
  dados: DadosPonto[]
  funcionarios: string[]
  mesAno: string
  erros: string[]
  debug: string[]
}

export default function PontoPage() {
  const [step, setStep] = useState<Step>("upload")
  const [relogioFile, setRelogioFile] = useState<File | null>(null)
  const [processing, setProcessing] = useState(false)
  const [previewData, setPreviewData] = useState<ProcessingResult | null>(null)
  const [finalWorkbook, setFinalWorkbook] = useState<XLSX.WorkBook | null>(null)
  const [showDebug, setShowDebug] = useState(false)

  const handleProcessar = useCallback(async () => {
    if (!relogioFile) return

    setProcessing(true)

    try {
      const relogioBuffer = await relogioFile.arrayBuffer()
      const relogioWorkbook = XLSX.read(relogioBuffer, { type: "array" })

      const resultado = processarRelogio(relogioWorkbook)

      setPreviewData(resultado)
      setStep("preview")
    } catch (error) {
      console.error("[v0] Erro ao processar:", error)
      setPreviewData({
        dados: [],
        funcionarios: [],
        mesAno: "",
        erros: [`Erro ao processar arquivo: ${error instanceof Error ? error.message : "Erro desconhecido"}`],
        debug: [],
      })
      setStep("preview")
    }

    setProcessing(false)
  }, [relogioFile])

  const handleGerarTabela = useCallback(() => {
    if (!previewData || previewData.dados.length === 0) return

    setProcessing(true)

    try {
      const workbook = gerarTabelaDados(previewData.dados, previewData.mesAno)
      setFinalWorkbook(workbook)
      setStep("result")
    } catch (error) {
      console.error("[v0] Erro ao gerar tabela:", error)
    }

    setProcessing(false)
  }, [previewData])

  const handleDownload = useCallback(() => {
    if (!finalWorkbook) return

    const wbout = XLSX.write(finalWorkbook, { bookType: "xlsx", type: "array" })
    const blob = new Blob([wbout], { type: "application/octet-stream" })
    const url = URL.createObjectURL(blob)

    const a = document.createElement("a")
    a.href = url
    a.download = `Dados_Ponto_${previewData?.mesAno?.replace("/", "-") || "extraidos"}.xlsx`
    document.body.appendChild(a)
    a.click()
    document.body.removeChild(a)
    URL.revokeObjectURL(url)
  }, [finalWorkbook, previewData?.mesAno])

  const handleReset = useCallback(() => {
    setStep("upload")
    setRelogioFile(null)
    setPreviewData(null)
    setFinalWorkbook(null)
    setShowDebug(false)
  }, [])

  return (
    <main className="min-h-screen bg-background px-4 py-8">
      <div className="mx-auto max-w-3xl">
        {/* Header */}
        <div className="mb-8 text-center">
          <div className="mb-4 flex items-center justify-center gap-3">
            <div className="flex h-12 w-12 items-center justify-center rounded-xl bg-primary">
              <Clock className="h-6 w-6 text-primary-foreground" />
            </div>
            <h1 className="text-2xl font-bold text-foreground">Extrator de Ponto</h1>
          </div>
          <p className="text-muted-foreground">
            Extraia os dados do relógio e gere uma planilha completa com cálculo de horas extras
          </p>
        </div>

        {/* Progress Steps */}
        <div className="mb-8 flex items-center justify-center gap-2">
          {["Upload", "Verificar", "Download"].map((label, idx) => {
            const stepOrder = ["upload", "preview", "result"]
            const currentIdx = stepOrder.indexOf(step)
            const isActive = idx === currentIdx
            const isCompleted = idx < currentIdx

            return (
              <div key={label} className="flex items-center gap-2">
                <div
                  className={`flex h-8 w-8 items-center justify-center rounded-full text-sm font-medium transition-colors ${
                    isActive
                      ? "bg-primary text-primary-foreground"
                      : isCompleted
                        ? "bg-emerald-500 text-white"
                        : "bg-muted text-muted-foreground"
                  }`}
                >
                  {isCompleted ? <CheckCircle2 className="h-4 w-4" /> : idx + 1}
                </div>
                <span
                  className={`text-sm ${isActive ? "font-medium text-foreground" : "text-muted-foreground"}`}
                >
                  {label}
                </span>
                {idx < 2 && <div className="mx-2 h-px w-8 bg-border" />}
              </div>
            )
          })}
        </div>

        {/* Step: Upload */}
        {step === "upload" && (
          <Card>
            <CardHeader>
              <CardTitle className="flex items-center gap-2">
                <FileSpreadsheet className="h-5 w-5" />
                Selecione o Arquivo do Relógio
              </CardTitle>
              <CardDescription>
                Arraste ou clique para adicionar o arquivo exportado do sistema de ponto eletrônico
              </CardDescription>
            </CardHeader>
            <CardContent className="flex flex-col gap-6">
              <FileUpload
                label="Logs de Comparecimento (Relógio)"
                description="Arquivo Excel (.xlsx ou .xls) exportado do sistema de ponto"
                file={relogioFile}
                onFileChange={setRelogioFile}
                status={relogioFile ? "success" : "idle"}
              />

              {/* Info Box */}
              <div className="rounded-lg border border-blue-200 bg-blue-50 p-4 dark:border-blue-900 dark:bg-blue-950">
                <div className="flex items-start gap-2">
                  <Info className="mt-0.5 h-5 w-5 flex-shrink-0 text-blue-600 dark:text-blue-400" />
                  <div className="text-sm text-blue-800 dark:text-blue-200">
                    <p className="font-medium">Como funciona:</p>
                    <ol className="mt-2 list-inside list-decimal space-y-1 text-blue-700 dark:text-blue-300">
                      <li>Faça upload do arquivo do relógio de ponto</li>
                      <li>O sistema extrai e calcula horas trabalhadas/extras</li>
                      <li>Baixe a planilha completa com alertas e observações</li>
                    </ol>
                  </div>
                </div>
              </div>

              <Button
                onClick={handleProcessar}
                disabled={!relogioFile || processing}
                className="w-full"
                size="lg"
              >
                {processing ? (
                  <>
                    <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                    Processando...
                  </>
                ) : (
                  "Processar Arquivo"
                )}
              </Button>
            </CardContent>
          </Card>
        )}

        {/* Step: Preview */}
        {step === "preview" && previewData && (
          <Card>
            <CardHeader>
              <CardTitle className="flex items-center gap-2">
                {previewData.dados.length > 0 ? (
                  <CheckCircle2 className="h-5 w-5 text-emerald-500" />
                ) : (
                  <AlertCircle className="h-5 w-5 text-amber-500" />
                )}
                Dados Encontrados
              </CardTitle>
              <CardDescription>
                Verifique se os dados estão corretos antes de gerar a tabela
              </CardDescription>
            </CardHeader>
            <CardContent className="flex flex-col gap-6">
              {/* Stats */}
              <div className="grid grid-cols-2 gap-4">
                <div className="flex items-center gap-3 rounded-lg bg-muted p-4">
                  <Users className="h-8 w-8 text-primary" />
                  <div>
                    <p className="text-2xl font-bold">{previewData.funcionarios.length}</p>
                    <p className="text-sm text-muted-foreground">Funcionários</p>
                  </div>
                </div>
                <div className="flex items-center gap-3 rounded-lg bg-muted p-4">
                  <Calendar className="h-8 w-8 text-primary" />
                  <div>
                    <p className="text-2xl font-bold">{previewData.dados.length}</p>
                    <p className="text-sm text-muted-foreground">Registros</p>
                  </div>
                </div>
              </div>

              {/* Funcionários */}
              {previewData.funcionarios.length > 0 && (
                <div>
                  <h3 className="mb-2 text-sm font-medium">Funcionários encontrados:</h3>
                  <div className="flex flex-wrap gap-2">
                    {previewData.funcionarios.map((nome) => (
                      <span key={nome} className="rounded-full bg-primary/10 px-3 py-1 text-sm text-primary">
                        {nome}
                      </span>
                    ))}
                  </div>
                </div>
              )}

              {/* Preview dos dados */}
              {previewData.dados.length > 0 && (
                <div>
                  <h3 className="mb-2 text-sm font-medium">Amostra dos dados:</h3>
                  <div className="max-h-48 overflow-auto rounded-lg border">
                    <table className="w-full text-sm">
                      <thead className="bg-muted">
                        <tr>
                          <th className="px-3 py-2 text-left font-medium">Funcionário</th>
                          <th className="px-3 py-2 text-left font-medium">Dia</th>
                          <th className="px-3 py-2 text-left font-medium">Entrada</th>
                          <th className="px-3 py-2 text-left font-medium">Saída</th>
                          <th className="px-3 py-2 text-left font-medium">Entrada</th>
                          <th className="px-3 py-2 text-left font-medium">Saída</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y">
                        {previewData.dados.slice(0, 10).map((d, i) => (
                          <tr key={i} className="hover:bg-muted/50">
                            <td className="px-3 py-2">{d.nome}</td>
                            <td className="px-3 py-2">{d.dia}</td>
                            <td className="px-3 py-2">{d.entrada1 || "-"}</td>
                            <td className="px-3 py-2">{d.saida1 || "-"}</td>
                            <td className="px-3 py-2">{d.entrada2 || "-"}</td>
                            <td className="px-3 py-2">{d.saida2 || "-"}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                  {previewData.dados.length > 10 && (
                    <p className="mt-2 text-xs text-muted-foreground">
                      Mostrando 10 de {previewData.dados.length} registros
                    </p>
                  )}
                </div>
              )}

              {/* Avisos */}
              {previewData.erros.length > 0 && (
                <div className="rounded-lg border border-amber-200 bg-amber-50 p-4 dark:border-amber-900 dark:bg-amber-950">
                  <div className="flex items-start gap-2">
                    <AlertCircle className="mt-0.5 h-5 w-5 text-amber-600 dark:text-amber-400" />
                    <div>
                      <p className="font-medium text-amber-800 dark:text-amber-200">Avisos</p>
                      <ul className="mt-1 text-sm text-amber-700 dark:text-amber-300">
                        {previewData.erros.map((erro, i) => (
                          <li key={i}>{erro}</li>
                        ))}
                      </ul>
                    </div>
                  </div>
                </div>
              )}

              {/* Registros incompletos */}
              {previewData.dados.filter((d) => d.status.includes("Incompleto")).length > 0 && (
                <div className="rounded-lg border border-amber-200 bg-amber-50 p-4 dark:border-amber-900 dark:bg-amber-950">
                  <p className="mb-2 font-medium text-amber-800 dark:text-amber-200">
                    Registros incompletos (batidas faltando):
                  </p>
                  <div className="max-h-32 overflow-y-auto">
                    <ul className="space-y-1 text-sm text-amber-700 dark:text-amber-300">
                      {previewData.dados
                        .filter((d) => d.status.includes("Incompleto"))
                        .slice(0, 10)
                        .map((d, i) => (
                          <li key={i}>
                            {d.nome} - dia {d.dia}: {d.status}
                          </li>
                        ))}
                      {previewData.dados.filter((d) => d.status.includes("Incompleto")).length > 10 && (
                        <li className="font-medium">
                          ... e mais {previewData.dados.filter((d) => d.status.includes("Incompleto")).length - 10} registros
                        </li>
                      )}
                    </ul>
                  </div>
                </div>
              )}

              {/* Info sobre funcionalidades */}
              <div className="rounded-lg border border-emerald-200 bg-emerald-50 p-4 dark:border-emerald-900 dark:bg-emerald-950">
                <div className="flex items-start gap-2">
                  <CheckCircle2 className="mt-0.5 h-5 w-5 flex-shrink-0 text-emerald-600 dark:text-emerald-400" />
                  <div className="text-sm text-emerald-800 dark:text-emerald-200">
                    <p className="font-medium">A planilha inteligente incluirá:</p>
                    <ul className="mt-2 list-inside list-disc space-y-1 text-emerald-700 dark:text-emerald-300">
                      <li>Fórmulas Excel para cálculos automáticos editáveis</li>
                      <li>Cores: verde (OK), vermelho (falta), amarelo (alerta), roxo (feriado)</li>
                      <li>Coluna de justificativas para faltas (atestado, férias, etc.)</li>
                      <li>Identificação automática de dias sem registro</li>
                      <li>Alertas de almoço menor que 1 hora</li>
                      <li>Resumo geral com totais de todos funcionários</li>
                    </ul>
                  </div>
                </div>
              </div>

              {/* Debug logs toggle */}
              {previewData.debug.length > 0 && (
                <div className="rounded-lg border bg-muted/50 p-4">
                  <button
                    onClick={() => setShowDebug(!showDebug)}
                    className="flex w-full items-center justify-between text-sm font-medium"
                  >
                    <span>Log de processamento (diagnóstico)</span>
                    {showDebug ? (
                      <ChevronUp className="h-4 w-4" />
                    ) : (
                      <ChevronDown className="h-4 w-4" />
                    )}
                  </button>
                  {showDebug && (
                    <div className="mt-3 max-h-60 overflow-y-auto rounded bg-background p-3">
                      <pre className="whitespace-pre-wrap font-mono text-xs text-muted-foreground">
                        {previewData.debug.join("\n")}
                      </pre>
                    </div>
                  )}
                </div>
              )}

              <div className="flex gap-3">
                <Button variant="outline" onClick={handleReset} className="flex-1">
                  Voltar
                </Button>
                <Button
                  onClick={handleGerarTabela}
                  disabled={processing || previewData.dados.length === 0}
                  className="flex-1"
                >
                  {processing ? (
                    <>
                      <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                      Gerando...
                    </>
                  ) : (
                    "Gerar Tabela de Dados"
                  )}
                </Button>
              </div>
            </CardContent>
          </Card>
        )}

        {/* Step: Result */}
        {step === "result" && (
          <Card>
            <CardHeader>
              <CardTitle className="flex items-center gap-2">
                <CheckCircle2 className="h-5 w-5 text-emerald-500" />
                Tabela Gerada com Sucesso
              </CardTitle>
              <CardDescription>Baixe a planilha inteligente com fórmulas e cores</CardDescription>
            </CardHeader>
            <CardContent className="flex flex-col gap-6">
              {/* Stats */}
              <div className="flex items-center justify-center gap-3 rounded-lg bg-emerald-50 p-6 dark:bg-emerald-950">
                <CheckCircle2 className="h-12 w-12 text-emerald-500" />
                <div className="text-center">
                  <p className="text-3xl font-bold text-emerald-700 dark:text-emerald-300">
                    {previewData?.dados.length || 0}
                  </p>
                  <p className="text-sm text-emerald-600 dark:text-emerald-400">registros extraídos</p>
                </div>
              </div>

              {/* Legenda de cores */}
              <div className="rounded-lg border border-blue-200 bg-blue-50 p-4 dark:border-blue-900 dark:bg-blue-950">
                <div className="flex items-start gap-2">
                  <Info className="mt-0.5 h-5 w-5 flex-shrink-0 text-blue-600 dark:text-blue-400" />
                  <div className="text-sm text-blue-800 dark:text-blue-200">
                    <p className="font-medium">Legenda de Cores na Planilha:</p>
                    <div className="mt-2 grid grid-cols-2 gap-2 text-blue-700 dark:text-blue-300">
                      <div className="flex items-center gap-2">
                        <span className="h-3 w-3 rounded bg-green-400"></span>
                        <span>Verde = Dia OK</span>
                      </div>
                      <div className="flex items-center gap-2">
                        <span className="h-3 w-3 rounded bg-red-400"></span>
                        <span>Vermelho = Sem registro</span>
                      </div>
                      <div className="flex items-center gap-2">
                        <span className="h-3 w-3 rounded bg-yellow-400"></span>
                        <span>Amarelo = Alerta/Incompleto</span>
                      </div>
                      <div className="flex items-center gap-2">
                        <span className="h-3 w-3 rounded bg-purple-400"></span>
                        <span>Roxo = Feriado/Fim de semana</span>
                      </div>
                    </div>
                  </div>
                </div>
              </div>

              {/* Funcionalidades */}
              <div className="rounded-lg bg-muted p-4">
                <h3 className="mb-2 font-medium">Funcionalidades da planilha inteligente:</h3>
                <ul className="list-inside list-disc space-y-1 text-sm text-muted-foreground">
                  <li>Fórmulas Excel editáveis (altere horários e os cálculos atualizam)</li>
                  <li>Coluna de justificativa para faltas (atestado, férias, folga, etc.)</li>
                  <li>Jornada configurável (padrão 8h, editável na célula H3)</li>
                  <li>Aba RESUMO com totais e referências às abas individuais</li>
                  <li>Campos para assinatura do funcionário e responsável</li>
                  <li>Aba Dados_PROCV para integração com outras planilhas</li>
                </ul>
              </div>

              <div className="flex gap-3">
                <Button variant="outline" onClick={handleReset} className="flex-1">
                  Processar Outro Arquivo
                </Button>
                <Button onClick={handleDownload} disabled={!finalWorkbook} className="flex-1 bg-emerald-600 hover:bg-emerald-700">
                  <Download className="mr-2 h-4 w-4" />
                  Baixar Planilha Inteligente
                </Button>
              </div>
            </CardContent>
          </Card>
        )}

        {/* Footer */}
        <p className="mt-8 text-center text-xs text-muted-foreground">
          O sistema detecta automaticamente o mês/ano e os funcionários a partir do arquivo do relógio.
        </p>
      </div>
    </main>
  )
}
