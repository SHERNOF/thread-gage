// document.addEventListener("DOMContentLoaded", () => update_model());
async function update_model(){
    let thread_type = document.getElementById('model-type').value

    let thread_gage_model = document.getElementById('model')
    try {
        const items = await eel.get_model(thread_type)()
        thread_gage_model.innerHTML = items.map((item) =>  
         `<option value="${item}">${item}</option>`   
        )
    }   catch (error) {
        thread_gage_model.innerHTML = `<option>Model</option>`;
    }
}

async function model_selected(){
    let selected = document.getElementById('model').value
    console.log(selected)
    let thread_type = document.getElementById('model-type').value
    console.log(thread_type)
    let model_title = document.getElementById('model-title')
    let myTable = document.getElementById('table')
    model_title.innerText = `Thread Gage Parameter of ${selected} as per AS1014`

    await eel.model_selected(selected, thread_type)()
    
}


