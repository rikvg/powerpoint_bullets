Attribute VB_Name = "Module1"
Sub ContextDots()
    On Error Resume Next
            With ActivePresentation
            
                SectionCount = .SectionProperties.Count
                
                For X = 1 To .Slides.Count
                    .Slides(X).Shapes("Background").Delete
                    .Slides(X).Shapes("Bullet").Delete
                    .Slides(X).Shapes("SectionTitleBox").Delete
                    
                    Set bg = .Slides(X).Shapes.AddShape(msoShapeRectangle, 0, 0, .PageSetup.SlideWidth, 28)
                    bg.Name = "Background"
                    ' Change the rectangle's colours here
                    bg.Fill.ForeColor.RGB = RGB(253, 187, 48)
                    ' bg.BorderWidth = 5 FIND OUT HOW THIS WORKS
                    'bg.Line.ForeColor.RGB = RGB(253, 187, 48)
                    
                    ' Change the bullets' size, shape and spacing here
                    BulletSize = 5
                    BulletShape = msoShapeOval
                    BulletSpacing = 2
                    
                    Offset = 20
                    SlideNumber = 1
                    For Y = 1 To SectionCount
                    
                        If Y <> 1 Then
                            Offset = (Y - 1) * .PageSetup.SlideWidth / (SectionCount)
                        End If
                        TextboxOffset = Offset - 7.5
                        
                        SectionSlidesCount = .SectionProperties.SlidesCount(Y)
                        sectionTitle = .SectionProperties.Name(Y)
                        Set textbox = .Slides(X).Shapes.AddTextbox(msoTextOrientationHorizontal, _
                               TextboxOffset, -2, 200, 50)
                        textbox.TextFrame.TextRange.Text = sectionTitle
                        textbox.Name = "SectionTitleBox"
                        
                        ' Change the font colour, size and type here
                        textbox.TextFrame.TextRange.Font.Color = vbBlack
                        textbox.TextFrame.TextRange.Font.Size = 9
                        textbox.TextFrame.TextRange.Font.Name = "Arial"
                        textbox.TextFrame.TextRange.Font.Bold = True
                        
                        For Z = 1 To SectionSlidesCount
                            Set Bullet = .Slides(X).Shapes.AddShape(BulletShape, _
                                Offset + (Z - 1) * (BulletSpacing + BulletSize), 16, BulletSize, BulletSize)
                            Bullet.Name = "Bullet"
                            Bullet.Line.Weight = 1
                            ' Change the bullets' fill and line colour here (case: Other slide)
                            Bullet.Fill.ForeColor.RGB = RGB(253, 187, 48)
                            Bullet.Line.ForeColor.RGB = vbBlack
                            
                            
                            If X = SlideNumber Then
                                ' textbox.TextFrame.TextRange.Font.Bold = True
                                ' Change the bullets' fill and line colour here (case: active slide)
                                Bullet.Fill.ForeColor.RGB = vbWhite
                            End If
                            
                            SlideNumber = SlideNumber + 1
                        Next Z:
                    Next Y:
                    
                Next X:
            End With
End Sub


