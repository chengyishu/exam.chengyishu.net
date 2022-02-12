// 正确答案
var answers = {};

$(function () {
    // 刷新、返回、关闭检测
    $(window).bind('beforeunload', function () {
        return '您所做的更改可能未保存。';
    });
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
                    var content = shuffle(data.content);
                    var mode = params.get('mode');
                    if (mode && mode == 'all') {
                        // 题海模式 (全部试题)
                    } else {
                        // 抽选模式 (抽选20题)
                        content = fetchTop(data.content, 20);
                    }
                    for (var [index, elem] of content.entries()) {
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
                            template += '<input name="q{{no}}" required class="form-control">';
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
                            template += '<textarea name="q{{no}}" required rows="1" class="form-control"></textarea>';
                            elem.answer = elem.comment;
                        } else if (elem.type == 'radio') {
                            // 单选题
                            elem.answer = elem.options[0];
                            for (var option of shuffle(elem.options)) {
                                template += '<div class="radio"><label><input type="radio" name="q{{no}}" required value="' + option + '"> <span>' + option + '</span></label></div>';
                            }
                        } else {
                            // 非法题型
                            $('#paper').append('<hr><div class="text-center">试卷题型有误 ...</div>');
                            return false;
                        }
                        // 追加答案
                        template += '<p class="answer q' + (index + 1) + '"><span class="tag">答案</span><span class="content">{{answer}}</span></p>';
                        template += '</div>';
                        // 内容替换
                        html = template.replaceAll('{{question}}', elem.question)
                            .replaceAll('{{score}}', elem.score)
                            .replaceAll('{{image}}', baseUrl + '/' + elem.image)
                            .replaceAll('{{audio}}', baseUrl + '/' + elem.audio)
                            .replaceAll('{{video}}', baseUrl + '/' + elem.video)
                            .replaceAll('{{no}}', index + 1)
                            .replaceAll('{{answer}}', elem.answer);
                        $('#paper').append(html);
                        // 满分信息
                        fullmarks += parseInt(elem.score);
                        // 正确答案
                        answers['q' + (index + 1)] = {
                            answer: elem.answer,
                            score: parseInt(elem.score),
                            type: elem.type,
                        };
                    }
                    // 满分信息
                    $('#fullmarks').append(fullmarks);
                    // 提交按钮
                    $('#paper').append('<br><hr><br><br><div class="row"><div class="col-xs-4"></div><div class="col-xs-4"><button type="submit" class="btn btn-primary btn-block">交卷打分</button></div><div class="col-xs-4"></div></div>');
                    $('#paper').on('submit', score);
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

// 获取数组前部元素
function fetchTop(array, size) {
    if (size >= array.length) {
        // 全部返回
        return array;
    }
    // 部分返回
    var parts = [];
    for (var i = 0; i < size; i++) {
        parts.push(array[i]);
    }
    return parts;
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
function score(e) {
    e.preventDefault();
    // 清除样式
    $('.form-group').removeClass('has-warning');
    $('.form-group').removeClass('has-error');
    // 关闭答案
    $('.answer').hide();
    // 重新打分
    var score = 0;
    for (var key in answers) {
        var write = undefined;
        if (answers[key].type == 'text') {
            // 填空题
            write = $('#paper input[name=' + key + ']');
        } else if (answers[key].type == 'textarea') {
            // 主观题
            write = $('#paper textarea[name=' + key + ']');
        } else if (answers[key].type == 'radio') {
            // 单选题
            write = $('#paper input[name=' + key + ']:checked');
        } else {
            // 非法题型
            return -1;
        }
        if (write.val() && write.val().trim() == answers[key].answer) {
            score += answers[key].score;
        } else {
            if (answers[key].type == 'textarea') {
                // 主观题采用警告样式
                $('.form-group').has(write).addClass('has-warning');
            } else {
                // 客观题采用错误样式
                $('.form-group').has(write).addClass('has-error');
            }
            // 显示答案
            $('.answer.' + key).show();
        }
    }
    $('#scorelabel').text('得分:');
    $('#score').text(score);
    $('#scoreboard').show();
    $('body,html').animate({scrollTop:0},100);
}