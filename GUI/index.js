// document.addEventListener("DOMContentLoaded", () => update_model());
async function update_model(){
    let thread_type = document.getElementById('model-type').value

    let thread_gage_model = document.getElementById('model')
    try {
        const items = await eel.get_model(thread_type)()
        console.log(items)
        thread_gage_model.innerHTML = items.map((item) =>  
         `<option value="${item}">${item}</option>`   
        )
    }   catch (error) {
        thread_gage_model.innerHTML = `<option>Model</option>`;
    }
}

async function model_selected(){
    let selected = document.getElementById('model').value
    let thread_type = document.getElementById('model-type').value
    let model_title = document.getElementById('model-title')
    let myTable = document.getElementById('table')
    model_title.innerText = `Thread Gage Parameter of ${selected} as per AS1014`
    let prm = document.getElementById('tarea')

    const params = await eel.model_selected(selected, thread_type)()
    console.log(params[3])

    let table = document.getElementById('table')
    table.innerHTML = 
    `<tr>
        <td>Pitch Diameter</td>
        <td>${params[3]}</td>
        <td>${params[4]}</td>
    </tr>`
    // let trP = document.createElement('tr')
    // let td1P = document.createElement('td')
    // let td2P = document.createElement('td')
    // let td3P = document.createElement('td')
    // let text1 = document.createTextNode('Pitch Diameter')
    // let text2 = document.createTextNode(String(params[3]))
    // let text3 = document.createTextNode(params[4])

    
    // td1P.appendChild(text1)
    // td2P.appendChild(text2)
    // td3P.appendChild(text3)
    // trP.appendChild(td1P)

    // table.appendChild(trP)

    
    // prm.innerHTML =  params.map((param) => 
    //     `<option>${param}</option> `
    // )
}


