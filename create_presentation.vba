
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
    para.Text = "Biological Inspiration and Artificial Neurons"
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
    para.Text = "Architecture of a Neural Network"
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
    para.Text = "Future Trends in Neural Networks"
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
    para.Text = "The core concept of interconnected nodes (neurons)."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Neural networks as function approximators."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "The power of learning from data."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    ' Add slide 4
    Set sld = ppt.Slides.Add(4, ppLayoutText)
    sld.Shapes.Title.TextFrame.TextRange.Text = "Biological Inspiration and Artificial Neurons"
    sld.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 600, 400)
    Set tf = shp.TextFrame
    tf.WordWrap = True
    tf.AutoSize = ppAutoSizeShapeToFitText

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Biological neurons and their functions."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "The structure of a biological neuron (dendrites, soma, axon)."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Simulating biological neurons with artificial neurons."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "The concept of weighted connections and activation functions."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "The role of synapses in biological and artificial networks."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    ' Add slide 5
    Set sld = ppt.Slides.Add(5, ppLayoutText)
    sld.Shapes.Title.TextFrame.TextRange.Text = "Architecture of a Neural Network"
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
    para.Text = "Output layer: Generating results."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Different network architectures (feedforward, recurrent, convolutional)."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Understanding the flow of information through layers."
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
    para.Text = "Feedforward Neural Networks (FNNs)."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Convolutional Neural Networks (CNNs) for image processing."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Recurrent Neural Networks (RNNs) for sequential data."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Long Short-Term Memory networks (LSTMs) for handling long sequences."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Autoencoders for dimensionality reduction and feature extraction."
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
    para.Text = "The role of cost functions and loss minimization."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Gradient descent optimization algorithm."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Backpropagation algorithm for adjusting weights."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Overfitting and regularization techniques."
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
    para.Text = "Image recognition and object detection."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Natural language processing (NLP)."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Speech recognition and synthesis."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Machine translation."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Medical diagnosis and drug discovery."
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
    para.Text = "Advantages:  High accuracy, adaptability, ability to handle complex data."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Disadvantages:  Computational cost, ""black box"" nature, need for large datasets."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Data dependency: Performance is highly dependent on the quality and quantity of data."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Overfitting and underfitting issues."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Interpretability challenges."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    ' Add slide 10
    Set sld = ppt.Slides.Add(10, ppLayoutText)
    sld.Shapes.Title.TextFrame.TextRange.Text = "Future Trends in Neural Networks"
    sld.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 600, 400)
    Set tf = shp.TextFrame
    tf.WordWrap = True
    tf.AutoSize = ppAutoSizeShapeToFitText

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Explainable AI (XAI) for increased transparency."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Spiking Neural Networks (SNNs) for energy efficiency."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Quantum neural networks for enhanced processing power."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Neuromorphic computing hardware."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Further advancements in deep learning techniques."
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
    para.Text = "Recap of the power and limitations of neural networks."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Discussion of the exciting future of neural network research."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Call to action: Encourage further exploration."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Q&A session."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

End Sub
