
// document.addEventListener("DOMContentLoaded",()=>{
    
// })

let thread_type = document.getElementById('model-type')
let thread_gage_model = document.getElementById('model')
let testing = load_type()
console.log(testing)

// getting the thread type
async function load_type(){
    let type = thread_type.value
    
     eel.thread_gage(type)()
    return type
}
console.log(load_type)

async function fetch_models(){
    let models = await eel.export_items()();
    // let models = if 
    console.log(models)
    models.forEach(model => {
        let option = document.createElement('option')
        option.text = model
        option.value = model
        thread_gage_model.add(option)
    })
}

// end of thread type

// async function send_type_to_python(){
//     // thread type to python
//     let type = thread_type.value
//     // await eel.thread_gage(type)
//     await eel.thread_gage(type)
// }

load_type()
// fetch_models()s
// send_type_to_python()