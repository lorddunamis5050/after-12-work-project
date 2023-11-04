const button = document.getElementById("change");
const text = document.getElementById("changetext");
let textisChanged = false


button.addEventListener("click", ()=>{
    if (!textisChanged){
        text.textContent = 'Text Changed'
    }else{
        text.textContent = 'Changing this text'
    }
    textisChanged = !textisChanged
})

