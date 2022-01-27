$(function () {
    // 加载题库
    var handle = setInterval(function () {
        if (window._book) {
            clearInterval(handle);
            var book = window._book;
            // 题库名称
            var bookName = window._bookName;
            $('#bookName').text(bookName);
            $.ajax({
                url: '/json/' + book + '/data.json',
                type: 'get',
                dataType: 'json',
                success: function (data) {
                    var template = '';
                    var html = '';
                    // 试卷列表
                    for (var elem of data) {
                        // 模板
                        template = '<a href="{{url}}" class="list-group-item"><h4 class="list-group-item-heading">{{title}}</h4><p class="list-group-item-text text-muted">{{description}}</p><a href="{{url}}&mode=all" class="btn btn-primary btn-sm float-right">题海模式</a></a>';
                        // 替换
                        html = template.replaceAll('{{url}}', '/app.html?book=' + book + '&paper=' + elem.id)
                            .replaceAll('{{title}}', elem.title)
                            .replaceAll('{{description}}', elem.description ? elem.description : '......');
                        // 渲染
                        $('#papers').append(html);
                    }
                }
            });
        }
    }, 1);
});
