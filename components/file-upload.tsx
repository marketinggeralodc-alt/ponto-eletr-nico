"use client"

import { useCallback, useMemo } from "react"
import { Upload, FileSpreadsheet, X, Check, AlertTriangle } from "lucide-react"
import { cn } from "@/lib/utils"

interface FileUploadProps {
  label: string
  description: string
  file: File | null
  onFileChange: (file: File | null) => void
  accept?: string
  status?: "idle" | "success" | "error"
}

function getFileExtension(filename: string): string {
  return filename.split(".").pop()?.toLowerCase() || ""
}

export function FileUpload({
  label,
  description,
  file,
  onFileChange,
  accept = ".xlsx,.xls,.csv",
  status = "idle",
}: FileUploadProps) {
  const fileWarning = useMemo(() => {
    if (!file) return null
    const ext = getFileExtension(file.name)
    if (ext === "pdf") {
      return "Arquivo PDF detectado. Por favor, use o arquivo Excel original (.xlsx ou .xls)"
    }
    if (!["xlsx", "xls", "csv"].includes(ext)) {
      return `Formato .${ext} pode não ser compatível. Recomendamos .xlsx ou .xls`
    }
    return null
  }, [file])

  const handleDrop = useCallback(
    (e: React.DragEvent<HTMLDivElement>) => {
      e.preventDefault()
      const droppedFile = e.dataTransfer.files[0]
      if (droppedFile) {
        onFileChange(droppedFile)
      }
    },
    [onFileChange]
  )

  const handleDragOver = useCallback((e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault()
  }, [])

  const handleFileInput = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const selectedFile = e.target.files?.[0]
      if (selectedFile) {
        onFileChange(selectedFile)
      }
    },
    [onFileChange]
  )

  const handleRemove = useCallback(() => {
    onFileChange(null)
  }, [onFileChange])

  return (
    <div className="flex flex-col gap-2">
      <label className="text-sm font-medium text-foreground">{label}</label>
      <div
        onDrop={handleDrop}
        onDragOver={handleDragOver}
        className={cn(
          "relative flex flex-col items-center justify-center rounded-lg border-2 border-dashed p-6 transition-colors",
          file
            ? status === "success"
              ? "border-emerald-500 bg-emerald-50"
              : status === "error"
                ? "border-red-500 bg-red-50"
                : "border-primary bg-primary/5"
            : "border-muted-foreground/25 hover:border-primary/50 hover:bg-muted/50"
        )}
      >
        {file ? (
          <div className="flex items-center gap-3">
            <div
              className={cn(
                "flex h-10 w-10 items-center justify-center rounded-full",
                status === "success" ? "bg-emerald-100" : status === "error" ? "bg-red-100" : "bg-primary/10"
              )}
            >
              {status === "success" ? (
                <Check className="h-5 w-5 text-emerald-600" />
              ) : (
                <FileSpreadsheet
                  className={cn("h-5 w-5", status === "error" ? "text-red-600" : "text-primary")}
                />
              )}
            </div>
            <div className="flex flex-col">
              <span className="text-sm font-medium text-foreground">{file.name}</span>
              <span className="text-xs text-muted-foreground">
                {(file.size / 1024).toFixed(1)} KB
              </span>
            </div>
            <button
              onClick={handleRemove}
              className="ml-2 rounded-full p-1 hover:bg-muted"
              type="button"
            >
              <X className="h-4 w-4 text-muted-foreground" />
            </button>
          </div>
        ) : (
          <>
            <Upload className="mb-2 h-8 w-8 text-muted-foreground" />
            <p className="text-sm text-muted-foreground">{description}</p>
            <p className="mt-1 text-xs text-muted-foreground/70">
              Arraste e solte ou clique para selecionar
            </p>
          </>
        )}
        <input
          type="file"
          accept={accept}
          onChange={handleFileInput}
          className="absolute inset-0 cursor-pointer opacity-0"
        />
      </div>
      {fileWarning && (
        <div className="flex items-center gap-2 rounded-md bg-amber-100 px-3 py-2 text-amber-800 dark:bg-amber-900 dark:text-amber-200">
          <AlertTriangle className="h-4 w-4 flex-shrink-0" />
          <span className="text-xs">{fileWarning}</span>
        </div>
      )}
    </div>
  )
}
