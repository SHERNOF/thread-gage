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
    
    `
     <tr>
        <th>Feature</th>
        <th>Min</th>
        <th>Max</th>
    </tr>
    <tr>
        <td class="td">Pitch Diameter</td>
        <td class="td">${params[3]}</td>
        <td class="td">${params[4]}</td>
    </tr>

     <tr>
        <td class="td">Pitch Diameter - Go</td>
        <td class="td">${params[5]}</td>
        <td class="td">${params[6]}</td>
    </tr>

     <tr>
        <td class="td">Pitch Diameter - NoGo</td>
        <td class="td">${params[5]}</td>
        <td class="td">${params[6]}</td>
    </tr>
    
    
    `

}


