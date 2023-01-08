myClickHandler() {
    var f = document.createElement('input');
    f.style.display='none';
    f.type='file';
    f.name='file';
    document.getElementById('yourformhere').appendChild(f);
    f.click();
}

myClickHandler