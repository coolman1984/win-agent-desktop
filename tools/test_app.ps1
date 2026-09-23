# A small real Windows Forms app for tools/smoke_windows.py: a text box, a checkbox, a
# drop-down list, a long list, a disabled button and a Submit button whose result shows
# in a label - the controls Notepad does not have.
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$f = New-Object System.Windows.Forms.Form
$f.Text = "wad test app"
$f.Width = 520; $f.Height = 460
$f.StartPosition = "Manual"; $f.Left = 120; $f.Top = 120

$tb = New-Object System.Windows.Forms.TextBox
$tb.Name = "nameBox"; $tb.AccessibleName = "Your name"
$tb.Left = 10; $tb.Top = 10; $tb.Width = 240

$cb = New-Object System.Windows.Forms.CheckBox
$cb.Name = "subscribe"; $cb.Text = "Subscribe"; $cb.Left = 10; $cb.Top = 45

$combo = New-Object System.Windows.Forms.ComboBox
$combo.Name = "color"; $combo.AccessibleName = "Color"; $combo.DropDownStyle = "DropDownList"
[void]$combo.Items.AddRange(@("Red", "Green", "Blue"))
$combo.Left = 10; $combo.Top = 75; $combo.Width = 160

$list = New-Object System.Windows.Forms.ListBox
$list.Name = "numbers"; $list.AccessibleName = "Numbers"
foreach ($i in 1..200) { [void]$list.Items.Add("item $i") }
$list.Left = 10; $list.Top = 110; $list.Width = 160; $list.Height = 120

$btn = New-Object System.Windows.Forms.Button
$btn.Name = "submit"; $btn.Text = "Submit"; $btn.Left = 10; $btn.Top = 245; $btn.Width = 120

$dead = New-Object System.Windows.Forms.Button
$dead.Name = "dead"; $dead.Text = "Disabled"; $dead.Enabled = $false
$dead.Left = 140; $dead.Top = 245; $dead.Width = 120

$lbl = New-Object System.Windows.Forms.Label
$lbl.Name = "status"; $lbl.Text = "waiting"; $lbl.Left = 10; $lbl.Top = 285; $lbl.Width = 480

$btn.Add_Click({ $lbl.Text = "Submitted: " + $tb.Text + " / " + $combo.SelectedItem + " / " + $cb.Checked })

$f.Controls.AddRange(@($tb, $cb, $combo, $list, $btn, $dead, $lbl))
[void]$f.ShowDialog()
