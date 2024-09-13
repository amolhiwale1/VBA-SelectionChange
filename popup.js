document.getElementById('submit-btn').addEventListener('click', function() {
    const sourceCells = document.getElementById('source-cells').value;
    const targetCells = document.getElementById('target-cells').value;

    // Convert the space-separated values to comma-separated values
    const formattedSource = sourceCells.trim().replace(/\s+/g, ', ');
    const formattedTarget = targetCells.trim().replace(/\s+/g, ', ');

    // Generate the VBA script using formatted values
    let vbaScript = "Public Sub Worksheet_SelectionChange(ByVal Target As Range)\n";
    const sourceArray = formattedSource.split(', ');
    const targetArray = formattedTarget.split(', ');

    for (let i = 0; i < sourceArray.length; i++) {
        vbaScript += `    If Not Intersect(Target, Me.Range("${sourceArray[i]}")) Is Nothing Then\n`;
        vbaScript += `        Application.Goto Me.Range("${targetArray[i]}"), True\n`;
        vbaScript += `    End If\n`;
    }
    vbaScript += "End Sub";

    // Display the generated script
    document.getElementById('vba-output').value = vbaScript;

    // Copy to clipboard on click
    document.getElementById('vba-output').addEventListener('click', function() {
        this.select();
        document.execCommand('copy');
        document.getElementById('copy-message').style.display = 'block';
        setTimeout(function() {
            document.getElementById('copy-message').style.display = 'none';
        }, 2000);
    });
});
