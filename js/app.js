$(function () {
    // 加载试卷
    const params = new URLSearchParams(window.location.search);
    var book = params.get('book');
    if (!book) {
        location.href = '/';
    } else {
        var paper = params.get('paper');
        if (!paper) {
            location.href = '/?book=' + book;
        } else {
            // 加载标题
            loadCache(book, paper);
            // 加载题目
            var baseUrl = '/json/' + book + '/' + paper;
            $.ajax({
                url: baseUrl + '/' + 'data.json',
                type: 'get',
                dataType: 'json',
                success: function (data) {
                    var template = '';
                    var html = '';
                    // 出题人和出题时间
                    $('#author').text(data.author);
                    $('#datetime').text(data.datetime);
                    // 满分信息
                    var fullmarks = 0;
                    // 试卷列表
                    for (var [index, elem] of shuffle(data.content).entries()) {
                        if (!elem.question) {
                            // 非法试卷
                            $('#paper').append('<hr><div class="text-center">试卷暂未开放 ...</div>');
                            return false;
                        }
                        // 模板编辑
                        template = '<div class="form-group"><label>{{no}}. {{question}}</label><span>（{{score}}分）</span>';
                        if (elem.image) {
                            // 图片附件
                            template += '<div class="image"><img src="{{image}}"></div>';
                        }
                        if (elem.audio) {
                            // 音频附件
                            template += '<div class="audio"><audio controls src="{{audio}}"></div>';
                        }
                        if (elem.video) {
                            // 视频附件
                            template += '<div class="video"><video controls src="{{video}}"></video></div>';
                        }
                        // 题目处理
                        if (elem.type == 'text') {
                            // 填空题
                            template += '<input name="q{{no}}" class="form-control">';
                            var match = elem.question.match(/\{(.*)\}/);
                            elem.answer = match[1];
                            var blank = ' ';
                            for (var i = 0; i < elem.answer.length * 3; i++) {
                                blank += '_';
                            }
                            blank += ' ';
                            elem.question = elem.question.replaceAll(match[0], blank);
                        } else if (elem.type == 'textarea') {
                            // 主观题
                            template += '<textarea name="q{{no}}" rows="1" class="form-control"></textarea>';
                            elem.answer = '参考: ' + elem.comment;
                        } else {
                            // 单选题
                            for (var option of shuffle(elem.options)) {
                                template += '<div class="radio"><label><input type="radio" name="q{{no}}" value="' + option + '"> ' + option + '</label></div>';
                            }
                            elem.answer = elem.options[0];
                        }
                        template += '</div>';
                        // 内容替换
                        html = template.replaceAll('{{question}}', elem.question)
                            .replaceAll('{{score}}', elem.score)
                            .replaceAll('{{image}}', baseUrl + '/' + elem.image)
                            .replaceAll('{{audio}}', baseUrl + '/' + elem.audio)
                            .replaceAll('{{video}}', baseUrl + '/' + elem.video)
                            .replaceAll('{{no}}', index + 1);
                        $('#paper').append(html);
                        // 满分信息
                        fullmarks += parseInt(elem.score);
                    }
                    // 满分信息
                    $('#fullmarks').append(fullmarks);
                    // 提交按钮
                    $('#paper').append('<hr><div class="row"><div class="col-xs-4"></div><div class="col-xs-4"><button type="button" class="btn btn-primary btn-block">提交</button></div><div class="col-xs-4"></div></div>');
                }
            });
        }
    }
});

// 打乱数组元素顺序
function shuffle(array) {
    let currentIndex = array.length, randomIndex;
    // While there remain elements to shuffle...
    while (currentIndex != 0) {
        // Pick a remaining element...
        randomIndex = Math.floor(Math.random() * currentIndex);
        currentIndex--;
        // And swap it with the current element.
        [array[currentIndex], array[randomIndex]] = [
            array[randomIndex], array[currentIndex]];
    }
    return array;
}

// 加载试卷标题信息
function loadCache(book, paper) {
    // 题库名称
    $.ajax({
        url: '/json/data.json',
        type: 'get',
        dataType: 'json',
        success: function (data) {
            for (var elem of data.all) {
                if (elem.id == book) {
                    $('#bookName').text(elem.title);
                    break;
                }
            }
        }
    });
    // 试卷名称
    $.ajax({
        url: '/json/' + book + '/data.json',
        type: 'get',
        dataType: 'json',
        success: function (data) {
            for (var elem of data) {
                if (elem.id == paper) {
                    $('#paperName').text(elem.title);
                    break;
                }
            }

        }
    });
}

// 计算分数
function score() {

}