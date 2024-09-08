import React, { useState, useEffect } from 'react'
import { Button } from "@/components/ui/button"
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card"
import { Progress } from "@/components/ui/progress"
import { FileSpreadsheet, BarChart2, Save } from "lucide-react"
import { AnalyzeExcel, SaveExcel, OpenFileDialog } from '../wailsjs/go/main/App'
import { EventsOn,EventsOff } from '../wailsjs/runtime'

export default function Component() { 
  const [filePath, setFilePath] = useState('')
  const [isAnalyzing, setIsAnalyzing] = useState(false)
  const [isAnalyzed, setIsAnalyzed] = useState(false)
  const [progress, setProgress] = useState({
    num:0,
    text:"初始化中..."
  })

  const handleFileSelect = async () => {
    try {
      const selectedFile = await OpenFileDialog()
      if (selectedFile) {
        setFilePath(selectedFile)
      }
    } catch (error) {
      console.error('File selection failed:', error)
    }
  }

  const handleAnalyze = async () => {
    if (!filePath) return
    setIsAnalyzing(true)
    setProgress({
      num:0,
      text:"初始化中..."
    })
    try {
      await AnalyzeExcel(filePath)
      setIsAnalyzed(true)
      alert('分析完成!')
    } catch (error) {
      console.error('Analysis failed:', error)
      alert('分析失败!')
    }
    setIsAnalyzing(false)
  }

  const handleSave = async () => {
    try {
      await SaveExcel()
      alert('保存成功!')
    } catch (error) {
      console.error('保存失败:', error)
    }
  }

  useEffect(() => {
    EventsOn('error', (error) => {
      alert(error)
    })
    EventsOn('progress', (progress) => {
      console.log(progress)
      setProgress(progress)
    })


    return () => {
      EventsOff('error')
      EventsOff('progress')
     
    }
  }, [isAnalyzing])

  return (
    <div className="min-h-screen flex items-center justify-center bg-gradient-to-br from-purple-400 via-pink-500 to-red-400">
      <Card className="w-full max-w-md shadow-2xl bg-white bg-opacity-90 backdrop-blur-sm border-4 border-transparent" style={{ borderImage: 'linear-gradient(to right, #6366f1, #ec4899) 1' }}>
        <CardHeader className="bg-gradient-to-r from-indigo-500 to-purple-600 text-white rounded-t-lg">
          <CardTitle className="text-3xl font-bold text-center">销售数据分析</CardTitle>
        </CardHeader>
        <CardContent className="mt-6 space-y-6 p-6">
          <Button onClick={handleFileSelect} className="w-full bg-gradient-to-r from-blue-500 to-cyan-500 hover:from-blue-600 hover:to-cyan-600 text-white shadow-lg transition-all duration-300">
            <FileSpreadsheet className="mr-2 h-5 w-5 text-blue-200" />
            选择原始数据文件
          </Button>
          {filePath && (
            //增加文字过长时自动换行
            <p className="text-sm text-gray-600 bg-gray-100 p-2 rounded-md text-wrap">
              所选文件:{filePath}
            </p>
          )}
          <Button 
            onClick={handleAnalyze} 
            disabled={!filePath || isAnalyzing} 
            className={`w-full shadow-lg transition-all duration-300 ${
              isAnalyzing 
                ? 'bg-gradient-to-r from-yellow-400 to-orange-500 hover:from-yellow-500 hover:to-orange-600' 
                : 'bg-gradient-to-r from-green-500 to-emerald-500 hover:from-green-600 hover:to-emerald-600'
            } text-white`}
          >
            <BarChart2 className="mr-2 h-5 w-5 text-green-200" />
            {isAnalyzing ? '分析中...' : '分析数据'}
          </Button>
          {isAnalyzing && (
            <div className="relative pt-1">
              <Progress value={progress.num} className="w-full h-2" />
              <div className="text-xs text-center mt-1 text-gray-600">{progress.num}%</div>
              <div className="text-xs text-center mt-1 text-gray-600">{progress.text}</div>
            </div>
          )}
          {isAnalyzed && (
            <Button onClick={handleSave} className="w-full bg-gradient-to-r from-pink-500 to-rose-500 hover:from-pink-600 hover:to-rose-600 text-white shadow-lg transition-all duration-300">
              <Save className="mr-2 h-5 w-5 text-pink-200" />
              保存分析好的文件
            </Button>
          )}
        </CardContent>
      </Card>
    </div>
  )
}