$('#result').on('click', '.row', function() {
    console.log('sanity check');
    $(this).find('li:nth-child(n+2)').toggle('display: none;');
});
