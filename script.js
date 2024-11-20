document.getElementById('documentForm').addEventListener('submit', function (event) {
    event.preventDefault();
    
    const title = document.getElementById('title').value;
    const author = document.getElementById('author').value;
    const university = document.getElementById('university').value;
    const year = document.getElementById('year').value;
    const content = document.getElementById('content').value;

    const data = {
        title: title,
        author: author,
        university: university,
        year: year,
        content: content
    };

    fetch('/generate-document', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(data)
    })
    .then(response => response.blob())
    .then(blob => {
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.style.display = 'none';
        a.href = url;
        a.download = 'document.docx';
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
    })
    .catch((error) => {
        console.error('Ошибка:', error);
    });
});
