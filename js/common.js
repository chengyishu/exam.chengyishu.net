$(function () {
    // 加载导航栏
    $.ajax({
        url: '/json/data.json',
        type: 'get',
        dataType: 'json',
        success: function (data) {
            var template = '';
            var html = '';
            // 选中题库
            const params = new URLSearchParams(window.location.search);
            var book = params.get('book');
            if (!book) {
                book = data.top[0].id;
            }
            // 全局变量
            window._book = book;
            // 常用题库
            var isTopActive = false;
            for (var elem of data.top) {
                template = '<li class="{{active}}"><a href="{{url}}">{{title}}</a></li>';
                html = template.replaceAll('{{active}}', book == elem.id ? 'active' : '')
                    .replaceAll('{{url}}', '/?book=' + elem.id)
                    .replaceAll('{{title}}', elem.title);
                $('#menu').append(html);
                // 判断选中常用题库
                if (book == elem.id) {
                    isTopActive = true;
                }
            }
            // 全部题库
            html = '<li class="dropdown' + (isTopActive ? '' : ' active') + '"><a href="/" class="dropdown-toggle" data-toggle="dropdown">全部题库 <b class="caret"></b></a><ul class="dropdown-menu" role="menu">';
            for (var elem of data.all) {
                if (elem.id) {
                    template = '<li class="{{active}}"><a href="{{url}}">{{title}}</a></li>';
                    html += template.replaceAll('{{active}}', book == elem.id ? 'active' : '')
                        .replaceAll('{{url}}', '/?book=' + elem.id)
                        .replaceAll('{{title}}', elem.title);
                    // 全局变量
                    if (book == elem.id) {
                        window._bookName = elem.title;
                    }
                } else {
                    // 分割线
                    html += '<li class="divider"></li>';
                }
            }
            html += '</ul></li>';
            $('#menu').append(html);
        }
    });
});