var wb = XLSX.utils.book_new();

wb.Props = {
    Title: "Result",
    Subject: "Result",
    Author: "PBM",
    CreatedDate: new Date()
};

wb.SheetNames.push("Result Sheet");
ws_data = [["Roll","Name","Marks"]];

function s2ab(s) { 
    var buf = new ArrayBuffer(s.length); //convert s to arrayBuffer
    var view = new Uint8Array(buf);  //create uint8array as viewer
    for (var i=0; i<s.length; i++) view[i] = s.charCodeAt(i) & 0xFF; //convert to octet
    return buf;    
}

document.getElementById("addBtn").addEventListener("click",(e)=>{
    e.preventDefault();
    const roll = document.getElementById("roll").value
    const name = document.getElementById("name").value
    const marks = document.getElementById("marks").value

    const data = [roll, name, marks]

    ws_data.push(data)

    document.getElementById("roll").value = ""
    document.getElementById("name").value = ""
    document.getElementById("marks").value = ""
})

document.getElementById("buildBtn").addEventListener("click",(e)=>{
    e.preventDefault();
    var ws = XLSX.utils.aoa_to_sheet(ws_data);
    wb.Sheets["Result Sheet"] = ws;

    var wbout = XLSX.write(wb, {bookType:'xlsx',  type: 'binary'});
    saveAs(new Blob([s2ab(wbout)],{type:"application/octet-stream"}), 'test.xlsx');
    console.log(ws_data,ws)
    ws_data = [["Roll","Name","Marks"]];
})

