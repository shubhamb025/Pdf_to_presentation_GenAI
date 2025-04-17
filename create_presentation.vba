
Sub CreatePresentation()
    Dim ppt As Presentation
    Dim sld As Slide
    Dim shp As Shape
    Dim tf As TextFrame
    Dim para As TextRange
    
    ' Create a new presentation
    Set ppt = Application.Presentations.Add

    ' Add title slide
    Set sld = ppt.Slides.Add(1, ppLayoutTitle)
    sld.Shapes.Title.TextFrame.TextRange.Text = "Unveiling the Power of Neural Networks: A Deep Dive"
    sld.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    
    ' Add creator name if provided
    If sld.Shapes.HasTitle Then
        Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 400, 600, 50)
        Set tf = shp.TextFrame
        tf.TextRange.Text = "Created by: "
        tf.TextRange.Font.Size = 14
        tf.TextRange.Font.Color.RGB = RGB(128, 128, 128)  ' Gray color
        tf.HorizontalAlignment = ppAlignCenter
    End If

    ' Add index slide
    Set sld = ppt.Slides.Add(2, ppLayoutText)
    sld.Shapes.Title.TextFrame.TextRange.Text = "Index"
    sld.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 600, 400)
    Set tf = shp.TextFrame
    tf.WordWrap = True
    tf.MarginLeft = 20
    tf.MarginRight = 20

    ' Add index content

    ' Add number
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "1."
    para.ParagraphFormat.Bullet.Visible = False
    para.Font.Bold = True
    para.Font.Size = 14
    para.ParagraphFormat.SpaceAfter = 0
    para.ParagraphFormat.SpaceBefore = 6
    
    ' Add title on next line
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Introduction to Neural Networks"
    para.ParagraphFormat.Bullet.Visible = False
    para.Font.Size = 14
    para.ParagraphFormat.LeftIndent = 20
    para.ParagraphFormat.SpaceAfter = 12
    para.ParagraphFormat.SpaceBefore = 0

    ' Add number
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "2."
    para.ParagraphFormat.Bullet.Visible = False
    para.Font.Bold = True
    para.Font.Size = 14
    para.ParagraphFormat.SpaceAfter = 0
    para.ParagraphFormat.SpaceBefore = 6
    
    ' Add title on next line
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "The Biological Inspiration: Neurons and Synapses"
    para.ParagraphFormat.Bullet.Visible = False
    para.Font.Size = 14
    para.ParagraphFormat.LeftIndent = 20
    para.ParagraphFormat.SpaceAfter = 12
    para.ParagraphFormat.SpaceBefore = 0

    ' Add number
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "3."
    para.ParagraphFormat.Bullet.Visible = False
    para.Font.Bold = True
    para.Font.Size = 14
    para.ParagraphFormat.SpaceAfter = 0
    para.ParagraphFormat.SpaceBefore = 6
    
    ' Add title on next line
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Architecture of a Neural Network: Layers and Connections"
    para.ParagraphFormat.Bullet.Visible = False
    para.Font.Size = 14
    para.ParagraphFormat.LeftIndent = 20
    para.ParagraphFormat.SpaceAfter = 12
    para.ParagraphFormat.SpaceBefore = 0

    ' Add number
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "4."
    para.ParagraphFormat.Bullet.Visible = False
    para.Font.Bold = True
    para.Font.Size = 14
    para.ParagraphFormat.SpaceAfter = 0
    para.ParagraphFormat.SpaceBefore = 6
    
    ' Add title on next line
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Types of Neural Networks"
    para.ParagraphFormat.Bullet.Visible = False
    para.Font.Size = 14
    para.ParagraphFormat.LeftIndent = 20
    para.ParagraphFormat.SpaceAfter = 12
    para.ParagraphFormat.SpaceBefore = 0

    ' Add number
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "5."
    para.ParagraphFormat.Bullet.Visible = False
    para.Font.Bold = True
    para.Font.Size = 14
    para.ParagraphFormat.SpaceAfter = 0
    para.ParagraphFormat.SpaceBefore = 6
    
    ' Add title on next line
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Training a Neural Network: Backpropagation"
    para.ParagraphFormat.Bullet.Visible = False
    para.Font.Size = 14
    para.ParagraphFormat.LeftIndent = 20
    para.ParagraphFormat.SpaceAfter = 12
    para.ParagraphFormat.SpaceBefore = 0

    ' Add number
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "6."
    para.ParagraphFormat.Bullet.Visible = False
    para.Font.Bold = True
    para.Font.Size = 14
    para.ParagraphFormat.SpaceAfter = 0
    para.ParagraphFormat.SpaceBefore = 6
    
    ' Add title on next line
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Applications of Neural Networks"
    para.ParagraphFormat.Bullet.Visible = False
    para.Font.Size = 14
    para.ParagraphFormat.LeftIndent = 20
    para.ParagraphFormat.SpaceAfter = 12
    para.ParagraphFormat.SpaceBefore = 0

    ' Add number
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "7."
    para.ParagraphFormat.Bullet.Visible = False
    para.Font.Bold = True
    para.Font.Size = 14
    para.ParagraphFormat.SpaceAfter = 0
    para.ParagraphFormat.SpaceBefore = 6
    
    ' Add title on next line
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Advantages and Disadvantages of Neural Networks"
    para.ParagraphFormat.Bullet.Visible = False
    para.Font.Size = 14
    para.ParagraphFormat.LeftIndent = 20
    para.ParagraphFormat.SpaceAfter = 12
    para.ParagraphFormat.SpaceBefore = 0

    ' Add number
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "8."
    para.ParagraphFormat.Bullet.Visible = False
    para.Font.Bold = True
    para.Font.Size = 14
    para.ParagraphFormat.SpaceAfter = 0
    para.ParagraphFormat.SpaceBefore = 6
    
    ' Add title on next line
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "The Future of Neural Networks"
    para.ParagraphFormat.Bullet.Visible = False
    para.Font.Size = 14
    para.ParagraphFormat.LeftIndent = 20
    para.ParagraphFormat.SpaceAfter = 12
    para.ParagraphFormat.SpaceBefore = 0

    ' Add number
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "9."
    para.ParagraphFormat.Bullet.Visible = False
    para.Font.Bold = True
    para.Font.Size = 14
    para.ParagraphFormat.SpaceAfter = 0
    para.ParagraphFormat.SpaceBefore = 6
    
    ' Add title on next line
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Conclusion: Key Takeaways and Future Directions"
    para.ParagraphFormat.Bullet.Visible = False
    para.Font.Size = 14
    para.ParagraphFormat.LeftIndent = 20
    para.ParagraphFormat.SpaceAfter = 12
    para.ParagraphFormat.SpaceBefore = 0

    ' Add blank line before conclusion
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = ""
    para.ParagraphFormat.SpaceAfter = 6
    
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Conclusion"
    para.ParagraphFormat.Bullet.Visible = False
    para.Font.Bold = True
    para.Font.Size = 14
    para.ParagraphFormat.SpaceBefore = 6

    ' Add slide 3
    Set sld = ppt.Slides.Add(3, ppLayoutText)
    sld.Shapes.Title.TextFrame.TextRange.Text = "Introduction to Neural Networks"
    sld.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 600, 400)
    Set tf = shp.TextFrame
    tf.WordWrap = True
    tf.AutoSize = ppAutoSizeShapeToFitText

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "What are Neural Networks?"
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "A brief history of neural networks."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "The concept of artificial intelligence and machine learning."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Neural networks as a subset of machine learning."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Simple analogy to explain the basic function."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    ' Add slide 4
    Set sld = ppt.Slides.Add(4, ppLayoutText)
    sld.Shapes.Title.TextFrame.TextRange.Text = "The Biological Inspiration: Neurons and Synapses"
    sld.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 600, 400)
    Set tf = shp.TextFrame
    tf.WordWrap = True
    tf.AutoSize = ppAutoSizeShapeToFitText

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Biological neurons and their function."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Synapses and signal transmission."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "The connection to artificial neural networks."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "How the brain's parallel processing inspires NN design."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Limitations of the biological analogy."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    ' Add slide 5
    Set sld = ppt.Slides.Add(5, ppLayoutText)
    sld.Shapes.Title.TextFrame.TextRange.Text = "Architecture of a Neural Network: Layers and Connections"
    sld.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 600, 400)
    Set tf = shp.TextFrame
    tf.WordWrap = True
    tf.AutoSize = ppAutoSizeShapeToFitText

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Input layer: Receiving data."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Hidden layers: Processing information."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Output layer: Producing results."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Weights and biases: Adjusting the network's behavior."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Activation functions: Introducing non-linearity."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    ' Add slide 6
    Set sld = ppt.Slides.Add(6, ppLayoutText)
    sld.Shapes.Title.TextFrame.TextRange.Text = "Types of Neural Networks"
    sld.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 600, 400)
    Set tf = shp.TextFrame
    tf.WordWrap = True
    tf.AutoSize = ppAutoSizeShapeToFitText

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Perceptrons: The simplest form."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Multilayer Perceptrons (MLPs): Deeper networks."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Convolutional Neural Networks (CNNs): For image processing."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Recurrent Neural Networks (RNNs): For sequential data."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Long Short-Term Memory (LSTM) networks:  Advanced RNNs for long sequences."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    ' Add slide 7
    Set sld = ppt.Slides.Add(7, ppLayoutText)
    sld.Shapes.Title.TextFrame.TextRange.Text = "Training a Neural Network: Backpropagation"
    sld.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 600, 400)
    Set tf = shp.TextFrame
    tf.WordWrap = True
    tf.AutoSize = ppAutoSizeShapeToFitText

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "The concept of supervised learning."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Forward propagation: Calculating the output."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Backpropagation: Calculating the error."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Gradient descent: Adjusting weights and biases."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Loss functions: Measuring the error."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    ' Add slide 8
    Set sld = ppt.Slides.Add(8, ppLayoutText)
    sld.Shapes.Title.TextFrame.TextRange.Text = "Applications of Neural Networks"
    sld.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 600, 400)
    Set tf = shp.TextFrame
    tf.WordWrap = True
    tf.AutoSize = ppAutoSizeShapeToFitText

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Image recognition and classification."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Natural language processing (NLP)."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Speech recognition."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Self-driving cars."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Medical diagnosis."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    ' Add slide 9
    Set sld = ppt.Slides.Add(9, ppLayoutText)
    sld.Shapes.Title.TextFrame.TextRange.Text = "Advantages and Disadvantages of Neural Networks"
    sld.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 600, 400)
    Set tf = shp.TextFrame
    tf.WordWrap = True
    tf.AutoSize = ppAutoSizeShapeToFitText

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Advantages: Powerful pattern recognition, adaptability, scalability."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Disadvantages: Black box nature, data dependency, computational cost, overfitting."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Addressing limitations through techniques like regularization and dropout."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "The need for large datasets."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Computational resource requirements."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    ' Add slide 10
    Set sld = ppt.Slides.Add(10, ppLayoutText)
    sld.Shapes.Title.TextFrame.TextRange.Text = "The Future of Neural Networks"
    sld.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 600, 400)
    Set tf = shp.TextFrame
    tf.WordWrap = True
    tf.AutoSize = ppAutoSizeShapeToFitText

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Advancements in deep learning."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "New architectures and algorithms."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Ethical considerations and responsible AI."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Potential breakthroughs in various fields."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Ongoing research and development."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    ' Add slide 11
    Set sld = ppt.Slides.Add(11, ppLayoutText)
    sld.Shapes.Title.TextFrame.TextRange.Text = "Conclusion: Key Takeaways and Future Directions"
    sld.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 600, 400)
    Set tf = shp.TextFrame
    tf.WordWrap = True
    tf.AutoSize = ppAutoSizeShapeToFitText

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Summary of key concepts covered."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Recap of the advantages and limitations."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Emphasis on the transformative potential of neural networks."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Call to action: Further exploration and learning."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Q&A session."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

End Sub
