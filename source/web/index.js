$(document).ready(function () {
    // the "href" attribute of the modal trigger must specify the modal ID that wants to be triggered
    $('.modal').modal({ dismissible: false });
});

document.querySelector("#runpython").addEventListener('click', async () => {
    $("#modal1").modal('open');
    let files = await eel.run_python()();
    $("#modal1").modal('close');
    console.log(files);

    files.forEach(element => {
        let li = document.createElement('li');
        li.innerHTML = element;
        li.classList.add("collection-item");
        document.querySelector("#collection").appendChild(li);
    });
});