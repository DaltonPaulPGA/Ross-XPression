' Template data
Dim i as integer
Dim published as xpBaseObject
Dim alwaysPublish() as string = {""}

For i = 0 to self.ObjectCount - 1
    Self.GetObject(i,published)
    published.ShowInTemplateData = false

    For Each name as string in alwaysPublish
        If published.Name = name
            published.ShowInTemplateData = true
        End If
    Next
Next