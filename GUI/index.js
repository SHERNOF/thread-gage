
let thread_type = document.getElementById('model-type')
let thread_gage_model = document.getElementById('model')
// Setting the of thread type
model_type = ['metric', 'imperial']
type = 'metric'

// getting the thread type
async function load_type(){
    // let dropdown = document.getElementById('model-type')
    types = model_type
    types.forEach(type => {
        let option = document.createElement('option')
        option.text = type
        option.value = type
        thread_type.add(option)
    })  
    // console.log(thread_type.value)
    // await eel.get_model(thread_type.value)
}

// async function fetch_models (){
//     let models = await eel.export_items()();
//     console.log(models)
//     models.forEach(model => {
//         let option = document.createElement('option')
//         option.text = model
//         option.value = model
//         thread_gage_model.add(option)
//     })
// }

// end of thread type

// async function send_type_to_python(){
//     // thread type to python
//     let type = thread_type.value
//     // await eel.thread_gage(type)
//     await eel.thread_gage(type)
// }

load_type()
// fetch_models()
// send_type_to_python()