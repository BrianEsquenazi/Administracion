VERSION 5.00
Begin VB.Form PrgCalculoVencidoPt 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "PrgCalculoVencidoPt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
            Select Case ZZTipo
                Case "Hoja"
                    ZZMeses = ""
                    ZZFecha = "  /  /    "
                    ZZMesesRevalida = ""
                    ZZFechaRevalida = "  /  /    "
                    ZZRevalida = ""
            
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Hoja"
                    ZSql = ZSql + " Where Hoja = " + "'" + ZZPartida + "'"
                    ZSql = ZSql + " and Producto = " + "'" + Terminado.Text + "'"
                    spHoja = ZSql
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                    If rstHoja.RecordCount > 0 Then
                        ZZRevalida = IIf(IsNull(rstHoja!Revalida), "0", rstHoja!Revalida)
                        ZZMesesRevalida = IIf(IsNull(rstHoja!MesesRevalida), "0", rstHoja!MesesRevalida)
                        ZZFechaRevalida = IIf(IsNull(rstHoja!FechaRevalida), "  /  /    ", rstHoja!FechaRevalida)
                        ZZFecha = rstHoja!Fecha
                        rstHoja.Close
                    End If
                    
                    spTerminado = "ConsultaTerminado " + "'" + Terminado.Text + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        ZZMeses = IIf(IsNull(rstTerminado!Vida), "", rstTerminado!Vida)
                        rstTerminado.Close
                    End If
                    
                    If Val(ZZMeses) <> 0 Then
                    
                        If Val(ZZRevalida) <> 0 Then
                        
                            WVida = Int(Val(ZZMesesRevalida) * 0.75)
                            WMes = Val(Mid$(ZZFechaRevalida, 4, 2))
                            WAno = Val(Right$(ZZFechaRevalida, 4))
                            
                                Else
                                
                            WVida = Int(Val(ZZMeses) * 0.75)
                            WMes = Val(Mid$(ZZFecha, 4, 2))
                            WAno = Val(Right$(ZZFecha, 4))
                                
                        End If
                        
                        For Ciclo = 1 To WVida
                            WMes = WMes + 1
                            If WMes > 12 Then
                                WAno = WAno + 1
                                WMes = 1
                            End If
                        Next Ciclo
                        ZMes = Str$(WMes)
                        ZAno = Str$(WAno)
                        Call Ceros(ZMes, 2)
                        Call Ceros(ZAno, 4)
                        ZZOrdVto = ZAno + ZMes + "01"
                        
                        If ZZOrdVto < ZZFechaActual Then
                            WMuestra.Row = Cicla
                            
                            WMuestra.Col = 1
                            WMuestra.CellBackColor = &HC0FFFF
                            
                            WMuestra.Col = 2
                            WMuestra.CellBackColor = &HC0FFFF
                            
                            WMuestra.Col = 3
                            WMuestra.CellBackColor = &HC0FFFF
                            
                            WMuestra.Col = 4
                            WMuestra.CellBackColor = &HC0FFFF
                            
                            WMuestra.Col = 5
                            WMuestra.CellBackColor = &HC0FFFF
                            
                            WMuestra.Col = 6
                            WMuestra.CellBackColor = &HC0FFFF
                            
                            WMuestra.Col = 7
                            WMuestra.CellBackColor = &HC0FFFF
                           
                            WMuestra.Col = 8
                            WMuestra.CellBackColor = &HC0FFFF
                            
                            WMuestra.Col = 9
                            WMuestra.CellBackColor = &HC0FFFF
                            
                        End If
                
                    End If
                    
                Case "Guia In"
                    
                    
                Case Else
                    
            End Select

