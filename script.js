const btn = document.querySelector(".js-btn")
btn.addEventListener('click' , async ()=>{

    const serviceCode = document.querySelector(".service_code").value;
    const name = document.querySelector(".name").value;
    const address = document.querySelector(".address").value;
    const division = document.querySelector(".division").value;
    const pharmaName = document.querySelector(".pharma_name").value;
    const brandName = document.querySelector(".brand_name").value;
    const productName = document.querySelector(".product_name").value;
    const serialNumber = document.querySelector(".serial_number").value;
    const issueProblem = document.querySelector(".issue_problem").value;
    const defectByTechnician = document.querySelector(".defect_by_technician").value;
    const defectedParts = document.querySelector(".defected_parts").value;
    const partsCost = document.querySelector(".parts_cost").value;
    const repairedDate = document.querySelector(".repaired_date").value;
    const challanNumber = document.querySelector(".challan_number").value;
    const courierDetails = document.querySelector(".courier_details").value;
    const fileName = document.querySelector(".file_name").value;

    const response = await fetch('template.docx')
    const data = await response.arrayBuffer();

    const zip = new PizZip(data);

    const doc = new window.docxtemplater().loadZip(zip);
    
    const placeholders = {
        service_code: serviceCode,
        name: name,
        address: address,
        division: division,
        pharma_name: pharmaName,
        brand_name: brandName,
        product_name: productName,
        serial_number: serialNumber,
        issue_problem: issueProblem,
        defect_by_technician: defectByTechnician,
        defected_parts: defectedParts,
        parts_cost: partsCost,
        repaired_date: repairedDate,
        challan_number: challanNumber,
        courier_details: courierDetails,
        file_name: fileName
    };


    doc.setData(placeholders);
    try{
        doc.render();
    }catch(error){
        console.error("error rendering",error.properties );
    }

    const output = doc.getZip().generate({
        type:"blob",
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    });

    const link = document.createElement("a");
    link.href=URL.createObjectURL(output);
    if(fileName===""){
        link.download="filled_template.docx"
    }else{
        link.download=`${fileName}.docx`;
    }
    link.click();
})



