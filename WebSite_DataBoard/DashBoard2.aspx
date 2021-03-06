﻿<%@ Page Title="" Language="C#" MasterPageFile="~/Site.master" AutoEventWireup="true" CodeFile="DashBoard2.aspx.cs" Inherits="DashBoard" %>

<asp:Content ID="Content1" ContentPlaceHolderID="HeadContent" Runat="Server">
    <!-- ECharts单文件引入 -->
    <script src="./Scripts/echarts.min.js" type="text/javascript"></script>
    <script src="Scripts/macarons.js" type="text/javascript"></script>
    <script src="Scripts/vintage.js" type="text/javascript"></script>
    <script src="Scripts/blue.js" type="text/javascript"></script>
    <script src="Scripts/echarts.min.js" type="text/javascript"></script>
    <script src="Scripts/jquery-1.4.1.min.js" type="text/javascript"></script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" Runat="Server">
    <asp:Button ID="Button1" runat="server" Text="Button" onclientclick="return setECharts();" />
    <input id="Text1" name="text" type="text" runat="server" style="display:none;"
            value="C:\Users\CN0YLBBC\Documents\CS CN E2E list - 01012018 - 19122018.XLSX"/>
    <table style="width: 100%;">
        <tr>
            <td>
                <!-- 为ECharts准备一个具备大小（宽高）的Dom -->
                <div id="divECharts1"style="height:500px;border:1px solid #ccc;padding:10px;"></div>
            </td>
            <td>
                <!-- 为ECharts准备一个具备大小（宽高）的Dom -->
                <div id="divECharts2"style="height:500px;border:1px solid #ccc;padding:10px;"></div>
            </td>            
        </tr>
        <tr>
            <td>
                <!-- 为ECharts准备一个具备大小（宽高）的Dom -->
                <div id="divECharts3"style="height:500px;border:1px solid #ccc;padding:10px;"></div>
            </td>
            <td>
                <!-- 为ECharts准备一个具备大小（宽高）的Dom -->
                <div id="divECharts4"style="height:500px;border:1px solid #ccc;padding:10px;"></div>
            </td>            
        </tr>
    </table>
    <!-- ECharts单文件引入 -->
    <script src="./Scripts/echarts.min.js" type="text/javascript"></script>
    <script src="Scripts/macarons.js" type="text/javascript"></script>
    <script src="Scripts/vintage.js" type="text/javascript"></script>
    <script type="text/javascript">
        function setECharts1() {
            // 基于准备好的dom，初始化echarts图表
            var myChart = echarts.init(document.getElementById('divECharts1'),'macarons');
            // 过渡---------------------
            myChart.showLoading({
                text: '正在努力的读取数据中...',    //loading
            });
            // ajax getting data...............       

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
            var filename =  document.getElementById("Text1").value;
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
        function setECharts() {    
       // 基于准备好的dom，初始化echarts图表
        var myChart = echarts.init(document.getElementById('divECharts2'),'macarons');
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

</asp:Content>

