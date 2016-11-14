# VBScript

```VBScript
Option Explicit

Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")

Dim dataFile 
Set dataFile = FSO.OpenTextFile("stock.dat", 2, True)

Dim symbolList
symbolList = array("600000", "300101")

Dim symbol
For each symbol in symbolList
    Dim IE

    Set IE      = WScript.CreateObject("InternetExplorer.Application")
    IE.Visible  = True
    
    Dim url
    url = "http://stockdata.stock.hexun.com/"& symbol & ".shtml"
    IE.Navigate url
    
    ' Wait for loading the whole page
    Do
        WScript.Sleep 10000
    Loop While IE.ReadyState <> 4
    
    Dim doc
    Set doc     = IE.Document
    
    Dim name
    name        = doc.GetElementById("quoteName").innerHTML
    
    Dim current
    current     = doc.GetElementById("q_current").innerHTML
    
    Dim upDownPrice
    upDownPrice = doc.GetElementById("q_updownprice").innerHTML
    
    Dim upDownRate
    upDownRate  = doc.GetElementById("q_upDownRate").innerHTML
    
    Dim open
    open        = doc.GetElementById("q_open").innerHTML
    
    Dim high
    high        = doc.GetElementById("q_high").innerHTML
    
    Dim turnover
    turnover    = doc.GetElementById("q_tv").innerHTML
    
    Dim profitRate
    profitRate  = doc.GetElementById("q_profitrate").innerHTML
    
    Dim i_p000001
    i_p000001   = doc.GetElementById("q_i_p000001").innerHTML
    
    Dim i_u000001
    i_u000001   = doc.GetElementById("q_i_u000001").innerHTML
    
    Dim preClose
    preClose    = doc.GetElementById("q_preclose").innerHTML
    
    Dim low
    low        = doc.GetElementById("q_low").innerHTML
    
    Dim change
    change    = doc.GetElementById("q_change").innerHTML
    
    Dim amp
    amp  = doc.GetElementById("q_amp").innerHTML
    
    Dim i_p399001
    i_p399001   = doc.GetElementById("q_i_p399001").innerHTML
    
    Dim i_u399001
    i_u399001   = doc.GetElementById("q_i_u399001").innerHTML
    
    Dim line
    line = join(array(symbol, name, current, upDownPrice, upDownRate, open, high, turnover, profitRate, i_p000001, i_u000001, _
                                                                    preClose, low, change, amp, i_p399001, i_u399001), ",")
    MsgBox(line)
    
    dataFile.WriteLine(line)
    
    IE.Quit()
    Set IE      = Nothing
Next

dataFile.Close()
Set FSO = Nothing
```
