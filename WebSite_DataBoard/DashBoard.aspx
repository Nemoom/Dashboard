<%@ Page Language="C#" AutoEventWireup="true" CodeFile="DashBoard.aspx.cs" Inherits="DashBoard" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <link href="~/Styles/Site.css" rel="stylesheet" type="text/css" />
    <link rel="shortcut icon" type="image/x-icon" href="img/ExhibitionInfoShell_Festo_32x32.png" />
    <script src="Scripts/macarons.js" type="text/javascript"></script>
    <script src="Scripts/vintage.js" type="text/javascript"></script>
    <script src="Scripts/FestoColorsN.js" type="text/javascript"></script>
    <script src="Scripts/FestoColor.js" type="text/javascript"></script>
    <script src="Scripts/echarts.min.js" type="text/javascript"></script>
    <script src="Scripts/jquery-1.4.1.min.js" type="text/javascript"></script>
</head>
<body>
    <form id="form1" runat="server">
    <div class="page">
        <div class="header">
            <div class="title">
                <h1>
                    E2E Dashboard
                </h1>
            </div>
            <div class="loginDisplay">
                <asp:LoginView ID="HeadLoginView" runat="server" EnableViewState="false">
                    <AnonymousTemplate>
                        [ <a href="~/Account/Login.aspx" id="HeadLoginStatus" runat="server">Log In</a>
                        ]
                    </AnonymousTemplate>
                    <LoggedInTemplate>
                        Welcome <span class="bold">
                            <asp:LoginName ID="HeadLoginName" runat="server" />
                        </span>! [
                        <asp:LoginStatus ID="HeadLoginStatus" runat="server" LogoutAction="Redirect" LogoutText="Log Out"
                            LogoutPageUrl="~/" />
                        ]
                    </LoggedInTemplate>
                </asp:LoginView>
            </div>
            <div class="clear hideSkiplink">
                <asp:Menu ID="NavigationMenu" runat="server" CssClass="menu" EnableViewState="false"
                    IncludeStyleBlock="false" Orientation="Horizontal">
                    <Items>
                        <asp:MenuItem NavigateUrl="~/Default.aspx" Text="Home" />
                        <asp:MenuItem NavigateUrl="~/DashBoard.aspx" Text="DashBoard" />
                        <asp:MenuItem NavigateUrl="~/Data.aspx" Text="Data" />
                        <asp:MenuItem NavigateUrl="~/About.aspx" Text="About" />
                    </Items>
                </asp:Menu>
            </div>
        </div>
        <div class="main">
            <input id="Text1" name="text" type="text" runat="server" style="display:none;"
            value="demo.XLSX"/>
            <asp:Button ID="Button1" runat="server" Text="解析" 
            onclientclick="return setECharts();" />
            <!-- 为ECharts准备一个具备大小（宽高）的Dom -->
            <table style="width: 100%;">
                <tr style="height: 50%;">
                    <td style="width: 50%;">
                        <!-- 为ECharts准备一个具备大小（宽高）的Dom -->
                        <div id="divECharts1" style="height: 500px; border: 1px solid #ccc; padding: 10px;">
                        </div>
                    </td>
                    <td style="width: 50%;">
                        <!-- 为ECharts准备一个具备大小（宽高）的Dom -->
                        <div id="divECharts2" style="height: 500px; border: 1px solid #ccc; padding: 10px;">
                        </div>
                    </td>
                </tr>
                <tr style="height: 50%;">
                    <td style="width: 50%;">
                        <!-- 为ECharts准备一个具备大小（宽高）的Dom -->
                        <div id="divECharts3" style="height: 500px; border: 1px solid #ccc; padding: 10px;">
                        </div>
                    </td>
                    <td style="width: 50%;">
                        <!-- 为ECharts准备一个具备大小（宽高）的Dom -->
                        <div id="divECharts4" style="height: 500px; border: 1px solid #ccc; padding: 10px;">
                        </div>
                    </td>
                </tr>
            </table>
        </div>
        <div class="clear">
            Copyright © Festo AG & Co. KG
        </div>
    </div>
    <!-- ECharts单文件引入 -->
    <script src="./Scripts/echarts.min.js" type="text/javascript"></script>
    <script src="Scripts/macarons.js" type="text/javascript"></script>
    <script src="Scripts/vintage.js" type="text/javascript"></script>
    <script src="Scripts/FestoColorsN.js" type="text/javascript"></script>
    <script src="Scripts/FestoColor.js" type="text/javascript"></script>
    <script type="text/javascript">
        // 基于准备好的dom，初始化echarts图表
        var myChart1 = echarts.init(document.getElementById('divECharts1'),'FestoColor');
        // 过渡---------------------
        myChart1.showLoading({
            text: '正在努力的读取数据中...',    //loading话术
        });
        // 基于准备好的dom，初始化echarts图表
        var myChart2 = echarts.init(document.getElementById('divECharts2'),'FestoColor');
        // 过渡---------------------
        myChart2.showLoading({
            text: '正在努力的读取数据中...',    //loading话术
        });
        // 基于准备好的dom，初始化echarts图表
        var myChart3 = echarts.init(document.getElementById('divECharts3'),'FestoColor');
        // 过渡---------------------
        myChart3.showLoading({
            text: '正在努力的读取数据中...',    //loading话术
        });
        // 基于准备好的dom，初始化echarts图表
        var myChart4 = echarts.init(document.getElementById('divECharts4'),'FestoColor');
        // 过渡---------------------
        myChart4.showLoading({
            text: '正在努力的读取数据中...',    //loading话术
        });
    </script>
    <script type="text/javascript">
        function setECharts1(filepath) {
            // 基于准备好的dom，初始化echarts图表
            var myChart = echarts.init(document.getElementById('divECharts1'),'FestoColor');
            // 过渡---------------------
            myChart.showLoading({
                text: '正在努力的读取数据中...',    //loading
            });
            // ajax getting data...............       

            // 图表使用-------------------
            var option = {
                title: {
                    text: '项目完成度统计',
                    x: 'center'
                },
                tooltip: {
                    trigger: 'item',
                    formatter: "{a} <br/>{b} : {c} ({d}%)"
                },
                legend: {
                    orient: 'vertical',
                    x: 'left',
                    data: []
                },
                series: [
                    {
                        clockWise:false,
                        name: '数量',
                        type: 'pie',
                        radius: [0, '30%'],
                        center: ['50%', '60%'],
                        label: {
                            normal: {
                                position: 'inner'
                            }
                        },
                        labelLine: {
                            normal: {
                                show: false
                            }
                        },                   
                        data: []
                    },
                    {
                        clockWise:false,
                        name: '数量',
                        type: 'pie',
                        radius: ['40%', '55%'],
                        center: ['50%', '60%'],
                        label: {
                            normal: {
                                formatter: '{a|{a}}{abg|}\n{hr|}\n  {b|{b}：}{c}  {per|{d}%}  ',
                                backgroundColor: '#eee',
                                borderColor: '#aaa',
                                borderWidth: 1,
                                borderRadius: 4,
                                rich: {
                                    a: {
                                        color: '#999',
                                        lineHeight: 22,
                                        align: 'center'
                                    },
                                    hr: {
                                        borderColor: '#aaa',
                                        width: '100%',
                                        borderWidth: 0.5,
                                        height: 0
                                    },
                                    b: {
                                        fontSize: 16,
                                        lineHeight: 33
                                    },
                                    per: {
                                        color: '#eee',
                                        backgroundColor: '#334455',
                                        padding: [2, 4],
                                        borderRadius: 2
                                    }
                                }
                            }
                        },
                        labelLine: {
                            normal: {
                                length1: 20
                            }
                        },    
                        data: []
                    }
                ]
            };

            // 为echarts对象加载数据 
            myChart.setOption(option); 

            var names = [];    //类别数组（实际用来盛放X轴坐标值）
            var values1 = [];    //inside
            var values2 = [];    //outside
            //var filename =  document.getElementById("Text1").value;
            var filename =  filepath;
            $.ajax({
                type: "post",
                async: true,            //异步请求（同步请求将会锁住浏览器，用户其他操作必须等待请求完成才可以执行）
                url: "Ajax/Handler.ashx",    //请求发送到TestServlet处
                data: {filename:filename},
                dataType: "json",        //返回数据形式为json
                success: function (result) {
                    //请求成功时执行该函数内容，result即为服务器返回的json对象
                    if (result) {
                        //标题列表
                        for (var i = 0; i < result.length; i++) {
                            names.push({ "name": result[i].name });
                        }
                        //饼数据和连线标题
                        for (var i = 0; i < 2; i++) {
                            values1.push({ "value": result[i].value, "name": result[i].name });
                        }
                        for (var i = 2; i < result.length; i++) {
                            values2.push({ "value": result[i].value, "name": result[i].name });
                        }
                        myChart.hideLoading();    //隐藏加载动画
                        myChart.setOption({        //加载数据图表
                            legend: {
                                data: names,
                            },
                            series: [
                            {
                                data: values1,
                            },
                            {
                                data: values2,
                            },
                            ]
                        });
                    }
                },
                error: function (errorMsg) {
                    //请求失败时执行该函数
                    alert("图表请求数据失败!");
                    myChart.hideLoading();
                }
            });

             /*窗口自适应，关键代码*/
	        setTimeout(function (){
	            window.onresize = function () {
	        	    myChart.resize();	        	
	            }
	        },200)
        }
    </script>
    <script type="text/javascript">
        function setECharts2(values) {    
        // 基于准备好的dom，初始化echarts图表
        var myChart = echarts.init(document.getElementById('divECharts1'),'FestoColor');
        // 过渡---------------------
        myChart.showLoading({
            text: '正在努力的读取数据中...',    //loading话术
        });
        // 图表使用-------------------
        var option = {
            title: {
                text: '项目完成进度一览',
                x: 'center'
            },
            tooltip: {
                trigger: 'item',
                formatter: "{a} <br/>{b} : {c} ({d}%)"
            },
            legend: {
                orient: 'vertical',
                x: 'left',
                data: []
            },
            series: [
                {
                    clockWise:false,
                    name: '数量',
                    type: 'pie',
                    radius: [0, '30%'],
                    center: ['50%', '60%'],
                    label: {
                        normal: {
                            position: 'inner'
                        }
                    },
                    labelLine: {
                        normal: {
                            show: false
                        }
                    },                   
                    data: []
                },
                {
                    clockWise:false,
                    name: '数量',
                    type: 'pie',
                    radius: ['40%', '55%'],
                    center: ['50%', '60%'],
                    label: {
                        normal: {
                            formatter: '{a|{a}}{abg|}\n{hr|}\n  {b|{b}：}{c|{c}}  {per|{d}%}  ',
                            backgroundColor: '#eee',
                            borderColor: '#aaa',
                            borderWidth: 1,
                            borderRadius: 4,                           
                            rich: {
                                a: {
                                    color: '#999',
                                    lineHeight: 22,
                                    align: 'center'
                                },                              
                                hr: {
                                    borderColor: '#aaa',
                                    width: '100%',
                                    borderWidth: 0.5,
                                    height: 0
                                },
                                b: {
                                    color: '#999',
                                    fontSize: 16,
                                    lineHeight: 33
                                },
                                c: {
                                    color: '#999',
                                    fontSize: 16,
                                    lineHeight: 33
                                },
                                per: {
                                    color: '#eee',
                                    backgroundColor: '#334455',
                                    padding: [2, 4],
                                    borderRadius: 2
                                }
                            }
                        }
                    },
                    labelLine: {
                        normal: {
                            length1: 20
                        }
                    },    
                    data: []
                }
            ]
        };

        // 为echarts对象加载数据 
        myChart.setOption(option); 

        // ajax getting data...............    
        var names = [];    //类别数组（实际用来盛放X轴坐标值）
        var values1 = [];    //inside
        var values2 = [];    //outside
        var filename = document.getElementById("Text1").value;

        $.ajax({
            type: "post",
            async: true,            //异步请求（同步请求将会锁住浏览器，用户其他操作必须等待请求完成才可以执行）
            url: "Ajax/Handler.ashx",    //请求发送到TestServlet处
            data: {filename:filename},
            dataType: "json",        //返回数据形式为json
            success: function (result) {
                //请求成功时执行该函数内容，result即为服务器返回的json对象
                if (result) {
                    //标题列表
                    for (var i = 0; i < result.length; i++) {
                        names.push({ "name": result[i].name });
                    }
                    //饼数据和连线标题
                    for (var i = 0; i < 2; i++) {
                        values1.push({ "value": result[i].value, "name": result[i].name });
                    }
                    for (var i = 2; i < result.length; i++) {
                        values2.push({ "value": result[i].value, "name": result[i].name });
                    }
                    myChart.hideLoading();    //隐藏加载动画
                    myChart.setOption({        //加载数据图表
                        legend: {
                            data: names,
                        },
                        series: [
                        {
                            data: values1,
                        },
                        {
                            data: values2,
                        },
                        ]
                    });
                }
            },
            error: function (errorMsg) {
                //请求失败时执行该函数
                alert("图表请求数据失败!");
                myChart.hideLoading();
            }
        });
        // 基于准备好的dom，初始化echarts图表
        var myChart2 = echarts.init(document.getElementById('divECharts2'),'FestoColor');
        // 过渡---------------------
        myChart2.showLoading({
            text: '正在努力的读取数据中...',    //loading话术
        });

        var dates_perMonth=values.split(',');
        var values_perMonth = [];   
        for (var i = 0; i < dates_perMonth.length; i++) {
                        values_perMonth.push({ "value": dates_perMonth[i].value});
                    }
        var option2 = {
            tooltip: {
                trigger: 'axis',
                axisPointer: {
                    type: 'cross',
                    crossStyle: {
                        color: '#999'
                    }
                }
            },
            toolbox: {
                feature: {
                    dataView: {show: true, readOnly: false},
                    magicType: {show: true, type: ['line', 'bar']},
                    restore: {show: true},
                    saveAsImage: {show: true}
                }
            },
            legend: {
                data:['蒸发量','降水量','平均温度']
            },
            xAxis: [
                {
                    type: 'category',
                    data: ['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月'],
                    axisPointer: {
                        type: 'shadow'
                    }
                }
            ],
            yAxis: [
                {
                    type: 'value',
                    name: '订单量',
                    min: 0,
                    max: 1000,
                    interval: 50,
                    axisLabel: {
                        formatter: '{value}'
                    }
                },
                {
                    type: 'value',
                    name: '温度',
                    min: 0,
                    max: 25,
                    interval: 5,
                    axisLabel: {
                        formatter: '{value} °C'
                    }
                }
            ],
            series: [
                {
                    name:'蒸发量',
                    type:'bar',
                    data:dates_perMonth
                },
                {
                    name:'降水量',
                    type:'bar',
                    data:[2.6, 5.9, 9.0, 26.4, 28.7, 70.7, 175.6, 182.2, 48.7, 18.8, 6.0, 2.3]
                },
                {
                    name:'平均温度',
                    type:'line',
                    yAxisIndex: 1,
                    data:[2.0, 2.2, 3.3, 4.5, 6.3, 10.2, 20.3, 23.4, 23.0, 16.5, 12.0, 6.2]
                }
            ]
        }
        myChart2.hideLoading();    //隐藏加载动画
        myChart2.setOption(option2);
         /*窗口自适应，关键代码*/
	    setTimeout(function (){
	        window.onresize = function () {
	        	myChart.resize();
                myChart2.resize();	        	
	        }
	    },200)
        return false;
       }

    </script>
    <script type="text/javascript">
        function setECharts() {    
       // 基于准备好的dom，初始化echarts图表
        var myChart = echarts.init(document.getElementById('divECharts3'),'FestoColor');
        // 过渡---------------------
        myChart.showLoading({
            text: '正在努力的读取数据中...',    //loading话术
        });
        // ajax getting data...............       

        // 图表使用-------------------
        var option = {
//**********饼图pie******************************************************************************************************************************
            title: {
                text: '项目完成进度一览',
                x: 'center'
            },
            tooltip: {
                trigger: 'item',
                formatter: "{a} <br/>{b} : {c} ({d}%)"
            },
            legend: {
                orient: 'vertical',
                x: 'left',
                data: []
            },
            series: [
                {
                    clockWise:false,
                    name: '数量',
                    type: 'pie',
                    radius: [0, '30%'],
                    center: ['50%', '60%'],
                    label: {
                        normal: {
                            position: 'inner'
                        }
                    },
                    labelLine: {
                        normal: {
                            show: false
                        }
                    },                   
                    data: []
                },
                {
                    clockWise:false,
                    name: '数量',
                    type: 'pie',
                    radius: ['40%', '55%'],
                    center: ['50%', '60%'],
                    label: {
                        normal: {
                            formatter: '{a|{a}}{abg|}\n{hr|}\n  {b|{b}：}{c}  {per|{d}%}  ',
                            backgroundColor: '#eee',
                            borderColor: '#aaa',
                            borderWidth: 1,
                            borderRadius: 4,
                            // shadowBlur:3,
                            // shadowOffsetX: 2,
                            // shadowOffsetY: 2,
                            // shadowColor: '#999',
                            // padding: [0, 7],
                            rich: {
                                a: {
                                    color: '#999',
                                    lineHeight: 22,
                                    align: 'center'
                                },
                                // abg: {
                                //     backgroundColor: '#333',
                                //     width: '100%',
                                //     align: 'right',
                                //     height: 22,
                                //     borderRadius: [4, 4, 0, 0]
                                // },
                                hr: {
                                    borderColor: '#aaa',
                                    width: '100%',
                                    borderWidth: 0.5,
                                    height: 0
                                },
                                b: {
                                    fontSize: 16,
                                    lineHeight: 33
                                },
                                per: {
                                    color: '#eee',
                                    backgroundColor: '#334455',
                                    padding: [2, 4],
                                    borderRadius: 2
                                }
                            }
                        }
                    },
                    labelLine: {
                        normal: {
                            length1: 20
                        }
                    },    
                    data: []
                }
            ]
//****************************************************************************************************************************************
        };

        // 为echarts对象加载数据 
        myChart.setOption(option); 

        var names = [];    //类别数组（实际用来盛放X轴坐标值）
        var values1 = [];    //inside
        var values2 = [];    //outside
        var filename = document.getElementById("Text1").value;

        $.ajax({
            type: "post",
            async: true,            //异步请求（同步请求将会锁住浏览器，用户其他操作必须等待请求完成才可以执行）
            url: "Ajax/Handler.ashx",    //请求发送到TestServlet处
            data: {filename:filename},
            dataType: "json",        //返回数据形式为json
            success: function (result) {
                //请求成功时执行该函数内容，result即为服务器返回的json对象
                if (result) {
                    //标题列表
                    for (var i = 0; i < result.length; i++) {
                        names.push({ "name": result[i].name });
                    }
                    //饼数据和连线标题
                    for (var i = 0; i < 2; i++) {
                        values1.push({ "value": result[i].value, "name": result[i].name });
                    }
                    for (var i = 2; i < result.length; i++) {
                        values2.push({ "value": result[i].value, "name": result[i].name });
                    }
                    myChart.hideLoading();    //隐藏加载动画
                    myChart.setOption({        //加载数据图表
                        legend: {
                            data: names,
                        },
                        series: [
                        {
                            data: values1,
                        },
                        {
                            data: values2,
                        },
                        ]
                    });
                }
            },
            error: function (errorMsg) {
                //请求失败时执行该函数
                alert("图表请求数据失败!");
                myChart.hideLoading();
            }
        });

         /*窗口自适应，关键代码*/
	    setTimeout(function (){
	        window.onresize = function () {
	        	myChart.resize();	        	
	        }
	    },200)
        return false;
       }

    </script>
    <script type="text/javascript">
        function setLoading() {    
        // 基于准备好的dom，初始化echarts图表
        var myChart1 = echarts.init(document.getElementById('divECharts1'),'FestoColorsN');
        // 过渡---------------------
        myChart1.showLoading({
            text: '正在努力的读取数据中...',    //loading话术
        });
        // 基于准备好的dom，初始化echarts图表
        var myChart2 = echarts.init(document.getElementById('divECharts2'),'FestoColorsN');
        // 过渡---------------------
        myChart2.showLoading({
            text: '正在努力的读取数据中...',    //loading话术
        });
        // 基于准备好的dom，初始化echarts图表
        var myChart3 = echarts.init(document.getElementById('divECharts3'),'FestoColorsN');
        // 过渡---------------------
        myChart3.showLoading({
            text: '正在努力的读取数据中...',    //loading话术
        });
        // 基于准备好的dom，初始化echarts图表
        var myChart4 = echarts.init(document.getElementById('divECharts4'),'FestoColorsN');
        // 过渡---------------------
        myChart4.showLoading({
            text: '正在努力的读取数据中...',    //loading话术
        });
        }
    </script>
    </form>
</body>
</html>
