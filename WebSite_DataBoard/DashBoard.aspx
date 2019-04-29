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
    <script src="Scripts/FestoColor_6.js" type="text/javascript"></script>
    <script src="Scripts/echarts.min.js" type="text/javascript"></script>
    <script src="Scripts/jquery-1.4.1.min.js" type="text/javascript"></script>
</head>
<body>
    <form id="form1" runat="server">
    <div class="page">
        <div class="header">
            <div class="title">
                <h1>
                    Transparent kit
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
            <!-- 为ECharts准备一个具备大小（宽高）的Dom -->
            <asp:Panel ID="Panel1" runat="server" Height="1050">
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
                            <div id="divECharts6" style="height: 500px; border: 1px solid #ccc; padding: 10px;">
                            </div>
                        </td>
                    </tr>
                </table>
            </asp:Panel>
            <asp:Panel ID="Panel2" runat="server" Height="1050">
                <table style="width: 100%;">
                    <tr style="height: 50%;">
                        <td style="width: 50%;">
                            <!-- 为ECharts准备一个具备大小（宽高）的Dom -->
                            <div id="divECharts5" style="height: 500px; border: 1px solid #ccc; padding: 10px;">
                            </div>
                        </td>
                        <td style="width: 50%;">
                            <!-- 为ECharts准备一个具备大小（宽高）的Dom -->
                            <div id="divECharts4" style="height: 500px; border: 1px solid #ccc; padding: 10px;">
                            </div>
                        </td>
                    </tr>
                    <tr style="height: 50%;">
                        <td style="width: 50%;">
                            <!-- 为ECharts准备一个具备大小（宽高）的Dom -->
                            <div id="divECharts7" style="height: 500px; border: 1px solid #ccc; padding: 10px;">
                            </div>
                        </td>
                        <td style="width: 50%;">
                            <!-- 为ECharts准备一个具备大小（宽高）的Dom -->
                            <div id="divECharts8" style="height: 500px; border: 1px solid #ccc; padding: 10px;">
                            </div>
                        </td>
                    </tr>
                </table>
            </asp:Panel>            
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
    <script src="Scripts/FestoColor_6.js" type="text/javascript"></script>
    <script type="text/javascript">
        // 基于准备好的dom，初始化echarts图表
        var myChart1 = echarts.init(document.getElementById('divECharts1'),'FestoColor_6');
        var myChart2 = echarts.init(document.getElementById('divECharts2'),'FestoColor_6');
        var myChart3 = echarts.init(document.getElementById('divECharts3'),'FestoColor_6');
        var myChart4 = echarts.init(document.getElementById('divECharts4'),'FestoColor_6');
        var myChart5 = echarts.init(document.getElementById('divECharts5'),'FestoColor_6');
        var myChart6 = echarts.init(document.getElementById('divECharts6'),'FestoColor_6');
        var myChart7 = echarts.init(document.getElementById('divECharts7'),'FestoColor_6');
        var myChart8 = echarts.init(document.getElementById('divECharts8'),'FestoColor_6');
        // 过渡---------------------
        myChart1.showLoading({
            text: '正在努力的读取数据中...',   
        });
        myChart2.showLoading({
            text: '正在努力的读取数据中...',   
        });
        
        myChart3.showLoading({
            text: '正在努力的读取数据中...',  
        });
        myChart4.showLoading({
            text: '正在努力的读取数据中...',   
        });
        myChart5.showLoading({
            text: '正在努力的读取数据中...',   
        });
        myChart6.showLoading({
            text: '正在努力的读取数据中...',   
        });
        myChart7.showLoading({
            text: '正在努力的读取数据中...',   
        });
        myChart8.showLoading({
            text: '正在努力的读取数据中...',   
        });
        var count = 0;
        var time_count;
//        time_count = setInterval(function () {
//                count++;
//                if (count%2==0) {
////                    document.getElementById('divECharts1').style.display="none";
////                    document.getElementById('divECharts2').style.display="none";
////                    document.getElementById('divECharts3').style.display="none";
////                    document.getElementById('divECharts6').style.display="none";
////                    document.getElementById('divECharts5').style.display="block";
////                    document.getElementById('divECharts4').style.display="block";
////                    document.getElementById('divECharts7').style.display="block";
////                    document.getElementById('divECharts8').style.display="block";
//                    document.getElementById('Panel2').style.display="none";
//                    document.getElementById('Panel1').style.display="block";
//                } 
//                else {
////                    document.getElementById('divECharts1').style.display="block";
////                    document.getElementById('divECharts2').style.display="block";
////                    document.getElementById('divECharts3').style.display="block";
////                    document.getElementById('divECharts6').style.display="block";
////                    document.getElementById('divECharts5').style.display="none";
////                    document.getElementById('divECharts4').style.display="none";
////                    document.getElementById('divECharts7').style.display="none";
////                    document.getElementById('divECharts8').style.display="none";
//                    document.getElementById('Panel1').style.display="none";
//                    document.getElementById('Panel2').style.display="block";
//                }                  
//            }, 10000);
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
        function setECharts2(values_ECharts1,values_ECharts2,values_ECharts3,values_ECharts4,values_ECharts5,values_ECharts6,values_ECharts7,values_ECharts8) {    
        //*****************************************************************************************************************************
        //**                                                      myChart1                                                           **
        //*****************************************************************************************************************************
        // 基于准备好的dom，初始化echarts图表
        var myChart1 = echarts.init(document.getElementById('divECharts1'),'FestoColor_6');
        // 过渡---------------------
        myChart1.showLoading({
            text: '正在努力的读取数据中...',    //loading话术
        });

        var var_Echarts1=values_ECharts1.split(';');

        var count_PO_soCreated_perMonth = var_Echarts1[0].split('!');  
        var count_CreatedDelay_perMonth = var_Echarts1[1].split('!'); 
        var PO_CreationMonitor_YTD = var_Echarts1[2].split('!');   
        var values_PO_CreationMonitor_YTD = [];    
        values_PO_CreationMonitor_YTD.push({ "value": PO_CreationMonitor_YTD[0], "name": '正常创建PO数(YTD)' });
        values_PO_CreationMonitor_YTD.push({ "value": PO_CreationMonitor_YTD[1], "name": '延期创建PO数(YTD)' });        

        // 图表使用-------------------
        var option1 = {
                        title: {
                            text: 'PO Creation Process Monitor',  
                            subtext: 'LT > 0.5(WD)',
                            x: 'center'                           
                        },
                        tooltip : {
                            trigger: 'axis'
                        },
                        toolbox: {
                            show : false,
                            y: 'bottom',
                            feature : {
                                mark : {show: true},
                                dataView : {show: true, readOnly: false},
                                magicType : {show: true, type: ['line', 'bar', 'stack', 'tiled']},
                                restore : {show: true},
                                saveAsImage : {show: true}
                            }
                        },
                        calculable : true,
                        legend: {
                            bottom: 10,
                            left: 'center',
                            data:['创建PO数','延期创建PO数','正常创建PO数(YTD)','延期创建PO数(YTD)']
                            
                        },
                        xAxis : [
                            {
                                type : 'category',
                                splitLine : {show : false},
                                data : ['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月']
                            }
                        ],
                        yAxis : [
                            {
                                type : 'value',
                                position: 'left'
                            }
                        ],
                        series : [
                            {
                                name:'创建PO数',
                                type:'bar',
                                data:count_PO_soCreated_perMonth
                            },
                            {
                                name:'延期创建PO数',
                                type:'bar',                                
                                data:count_CreatedDelay_perMonth
                            },                        

                            {
                                name:'PO Creation Monitor(YTD)',
                                type:'pie',
                                tooltip : {
                                    trigger: 'item',
                                    formatter: '{a} <br/>{b} : {c} ({d}%)'
                                },
                                center: ['70%',130],
                                radius : [0, 50],
                                itemStyle :　{
                                    normal : {
                                        labelLine : {
                                            length : 20
                                        }
                                    }
                                },
                                data:values_PO_CreationMonitor_YTD                                
                            }
                        ]
                    };         
       
        myChart1.hideLoading();    //隐藏加载动画
        myChart1.setOption(option1);
        //*****************************************************************************************************************************
        //**                                                      myChart2                                                           **
        //*****************************************************************************************************************************
        // 基于准备好的dom，初始化echarts图表
        var myChart2 = echarts.init(document.getElementById('divECharts2'),'FestoColor_6');
        // 过渡---------------------
        myChart2.showLoading({
            text: '正在努力的读取数据中...',    //loading话术
        });

        var var_Echarts2=values_ECharts2.split(';');

        var count_PO_poCreated_perMonth = var_Echarts2[0].split('!');  
        var count_ReleaseDelay_perMonth = var_Echarts2[1].split('!'); 
        var PO_Release_YTD = var_Echarts2[2].split('!');   
        var values_PO_Release_YTD = [];    
        values_PO_Release_YTD.push({ "value": PO_Release_YTD[0], "name": '按时释放PO数(YTD)' });
        values_PO_Release_YTD.push({ "value": PO_Release_YTD[1], "name": '延期释放PO数(YTD)' });        

        // 图表使用-------------------
        var option2 = {
                        title: {
                            text: 'PO Release Process Monitor', 
                            subtext: 'LT > 5(WD)',
                            x: 'center'
                        },
                        tooltip : {
                            trigger: 'axis'
                        },
                        toolbox: {
                            show : false,
                            y: 'bottom',
                            feature : {
                                mark : {show: true},
                                dataView : {show: true, readOnly: false},
                                magicType : {show: true, type: ['line', 'bar', 'stack', 'tiled']},
                                restore : {show: true},
                                saveAsImage : {show: true}
                            }
                        },
                        calculable : true,
                        legend: {
                            bottom: 10,
                            left: 'center',
                            data:['创建PO数','延期释放PO数','按时释放PO数(YTD)','延期释放PO数(YTD)']
                        },
                        xAxis : [
                            {
                                type : 'category',
                                splitLine : {show : false},
                                data : ['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月']
                            }
                        ],
                        yAxis : [
                            {
                                type : 'value',
                                position: 'left'

                            }
                        ],
                        series : [
                            {
                                name:'创建PO数',
                                type:'bar',
                                data:count_PO_poCreated_perMonth
                            },
                            {
                                name:'延期释放PO数',
                                type:'bar',                                
                                data:count_ReleaseDelay_perMonth
                            },                        

                            {
                                name:'PO Release Monitor(YTD)',
                                type:'pie',
                                tooltip : {
                                    trigger: 'item',
                                    formatter: '{a} <br/>{b} : {c} ({d}%)'
                                },
                                center: ['70%',130],
                                radius : [0, 50],
                                itemStyle :　{
                                    normal : {
                                        labelLine : {
                                            length : 20
                                        }
                                    }
                                },
                                data:values_PO_Release_YTD                                
                            }
                        ]
                    };         
       
        myChart2.hideLoading();    //隐藏加载动画
        myChart2.setOption(option2);
        //*****************************************************************************************************************************
        //**                                                      myChart3                                                           **
        //*****************************************************************************************************************************
        // 基于准备好的dom，初始化echarts图表
        var myChart3 = echarts.init(document.getElementById('divECharts3'),'FestoColor_6');
        // 过渡---------------------
        myChart3.showLoading({
            text: '正在努力的读取数据中...',    //loading话术
        });

        var var_Echarts3=values_ECharts3.split(';');

        var var_DCR = var_Echarts3[0].split('!');  
        var var_DCR_0400 = var_Echarts3[1].split('!'); 
        var var_DCR_0481 = var_Echarts3[2].split('!');
          
////        var values_DCR_inside = []; 
////        var values_DCR_outside = [];    
////        values_DCR_inside.push({ "value": str_DCR_inside[0], "name": '已完成' });
////        values_DCR_inside.push({ "value": str_DCR_inside[1], "name": '未完成' }); 
////        values_DCR_outside.push({ "value": str_DCR_outside[0], "name": '及时交付' });
////        values_DCR_outside.push({ "value": str_DCR_outside[1], "name": '延期交付' }); 
////        values_DCR_outside.push({ "value": str_DCR_outside[2], "name": '已超期' });
////        values_DCR_outside.push({ "value": str_DCR_outside[3], "name": '预计超期' }); 
////        values_DCR_outside.push({ "value": str_DCR_outside[4], "name": '正常' });

        // 图表使用-------------------
        var option3 = {
            title: {
                        text: 'DCR(Delivery Class Reliablity)', 
                        subtext: 'According to Quotation LT',
                        x: 'center'
                    },
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
                bottom: 10,
                left: 'center',
                data:['0400 DCR','0481 DCR','0400+0481 DCR']
            },
            xAxis: [
                {
                    type: 'category',
                    data: ['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月','YTD'],
                    axisPointer: {
                        type: 'shadow'
                    }
                }
            ],
            yAxis: [
                {
                    type: 'value',
                    name: '完成率',
                    min: 0,
                    max: 100,
//                    interval: 50,
                    axisLabel: {
                        formatter: '{value}%'
                    }
                },        
            ],
            series: [                
                {
                    name:'0400 DCR',
                    type:'bar',
                    label: {
                        normal: {
                            show: true,
                            position: 'top',
                            formatter: function(params){
                                if (params.value > 0) {
                                    return params.value;
                                } else {
                                    return '';
                                }
                            }
                        }
                    },
//                    markPoint : {
//                         data : [
//                            {
//                                name: '最优值',
//                                type: 'max'
//                            }
//                        ],
//                        animationThreshold :78
//                    },
                    markLine : {
                        data : [
                            {name: 'Target',yAxis: 78}
                        ]
                    },
                    data:var_DCR_0400       
                },
                {
                    name:'0481 DCR',
                    type:'bar',
                    label: {
                        normal: {
                            show: true,
                            position: 'top',
                            formatter: function(params){
                                if (params.value > 0) {
                                    return params.value;
                                } else {
                                    return '';
                                }
                            }
                        }
                    },
//                    markPoint : {
//                         data : [
//                            {
//                                name: '最优值',
//                                type: 'max'
//                            }
//                        ],
//                        animationThreshold :78
//                    },
                    markLine : {
                        data : [
                            {name: 'Target',yAxis: 78}
                        ]
                    },
                    data:var_DCR_0481
        
                },
                {
                    name:'0400+0481 DCR',
                    type:'bar',
                    label: {
                        normal: {
                            show: true,
                            position: 'top',
                            formatter: function(params){
                                if (params.value > 0) {
                                    return params.value;
                                } else {
                                    return '';
                                }
                            }
                        }
                    },
//                    markPoint : {
//                         data : [
//                            {
//                                name: '最优值',
//                                type: 'max'
//                            }
//                        ],
//                        animationThreshold :78
//                    },
                    markLine : {
                        data : [
                            {name: 'Target',yAxis: 78}
                        ]
                    },
                    data:var_DCR
                },
            ]
        }     
       
        myChart3.hideLoading();    //隐藏加载动画
        myChart3.setOption(option3);
        //*****************************************************************************************************************************
        //**                                                      myChart4                                                           **
        //*****************************************************************************************************************************
         // 基于准备好的dom，初始化echarts图表
        var myChart4 = echarts.init(document.getElementById('divECharts4'),'FestoColor_6');
        // 过渡---------------------
        myChart4.showLoading({
            text: '正在努力的读取数据中...',    //loading话术
        });

        var var_Echarts4=values_ECharts4.split(';');

        var count_PO_Finished_perMonth=var_Echarts4[0].split('!');
        var count_PO_ForecastFinished_perMonth = var_Echarts4[1].split('!');   
        var NetValue_PO_Finished_perMonth = var_Echarts4[2].split('!');   

        var count_PO_Finished_perMonth1=count_PO_Finished_perMonth.splice(0,count_PO_Finished_perMonth.length-1);
        var count_PO_ForecastFinished_perMonth1=count_PO_ForecastFinished_perMonth.splice(0,count_PO_ForecastFinished_perMonth.length-1);
        var NetValue_PO_Finished_perMonth1 = NetValue_PO_Finished_perMonth.splice(0,NetValue_PO_Finished_perMonth.length-1);
       
        var option4 = {
            title: {
                        text: 'Finish POs Monitor', 
                        x: 'center'
                    },
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
                bottom: 10,
                left: 'center',
                selected: {
                '预计完成量': false
                },
                data:['订单完成量','预计完成量','订单金额']
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
//                    min: 0,
//                    max: 1000,
//                    interval: 50,
                    axisLabel: {
                        formatter: '{value}'
                    }
                },
                {
                    type: 'value',
                    name: '金额',
//                    min: 0,
//                    max: 50000000,
//                    interval: 5000,
                    axisLabel: {
                        formatter: '{value} 元'
                    }
                }
            ],
            series: [
                {
                    name:'订单完成量',
                    type:'bar',
                    data:count_PO_Finished_perMonth1
                },
                {
                    name:'预计完成量',
                    type:'bar',
                    data:count_PO_ForecastFinished_perMonth1       
                },
                {
                    name:'订单金额',
                    type:'line',
                    yAxisIndex: 1,
                    data:NetValue_PO_Finished_perMonth1
        
                }
            ]
        }
        myChart4.hideLoading();    //隐藏加载动画
        myChart4.setOption(option4);
        //*****************************************************************************************************************************
        //**                                                      myChart5                                                           **
        //*****************************************************************************************************************************
         // 基于准备好的dom，初始化echarts图表
        var myChart5 = echarts.init(document.getElementById('divECharts5'),'FestoColor_6');
        // 过渡---------------------
        myChart5.showLoading({
            text: '正在努力的读取数据中...',    //loading话术
        });

        var var_Echarts5=values_ECharts5.split(';');

        var count_SO_soCreated_perMonth=var_Echarts5[0].split('!');
        var NetValue_SO_soCreated_perMonth = var_Echarts5[1].split('!'); 
        var count_SO_soCreated_perMonth1=count_SO_soCreated_perMonth.splice(0,count_SO_soCreated_perMonth.length-1);
        var NetValue_SO_soCreated_perMonth1 = NetValue_SO_soCreated_perMonth.splice(0,NetValue_SO_soCreated_perMonth.length-1);  
       
        var option5 = {
            title: {
                        text: 'Create SOs Monitor', 
                        x: 'center'
                    },
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
                bottom: 10,
                left: 'center',
                data:['新进订单量','订单金额']
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
//                    min: 0,
//                    max: 1000,
//                    interval: 50,
                    axisLabel: {
                        formatter: '{value}'
                    }
                },
                {
                    type: 'value',
                    name: '金额',
//                    min: 0,
//                    max: 50000000,
//                    interval: 5000,
                    axisLabel: {
                        formatter: '{value} 元'
                    }
                }
            ],
            series: [
                {
                    name:'新进订单量',
                    type:'bar',
                    data:count_SO_soCreated_perMonth1
                },                
                {
                    name:'订单金额',
                    type:'line',
                    yAxisIndex: 1,
                    data:NetValue_SO_soCreated_perMonth1
        
                }
            ]
        }
        myChart5.hideLoading();    //隐藏加载动画
        myChart5.setOption(option5);
        //*****************************************************************************************************************************
        //**                                                      myChart6                                                           **
        //*****************************************************************************************************************************
         // 基于准备好的dom，初始化echarts图表
        var myChart6 = echarts.init(document.getElementById('divECharts6'),'FestoColor_6');
        // 过渡---------------------
        myChart6.showLoading({
            text: '正在努力的读取数据中...',    //loading话术
        });
     
        var var_Echarts6=values_ECharts6.split(';');

        var var_LT = var_Echarts6[0].split('!');  
        var var_LT_0400 = var_Echarts6[1].split('!'); 
        var var_LT_0481 = var_Echarts6[2].split('!');
        var option6 = {
            title: {
                        text: 'Average LeadTime', 
                        x: 'center'
                    },
            tooltip: {
                trigger: 'axis',
                axisPointer: {
                    type: 'cross',
                    crossStyle: {
                        color: '#999'
                    }
                }
            },         
            legend: {
                bottom: 10,
                left: 'center',
                data:['0400 LT','0481 LT','0400+0481 LT'],
                selected: {
                '0481 LT': false, '0400+0481 LT': false
                }
            },
            xAxis: [
                {
                    type: 'category',
                    data: ['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月','YTD'],
                    axisPointer: {
                        type: 'shadow'
                    }
                }
            ],
            yAxis: [
                {
                    type: 'value',
                    name: 'Average LeadTime',
//                    min: 0,
//                    max: 1000,
//                    interval: 50,
                    axisLabel: {
                        formatter: '{value}CDS'
                    }
                },               
            ],
            series: [                 
                {
                    name:'0400 LT',
                    type:'bar',
                    label: {
                        normal: {
                            show: true,
                            position: 'top',
                            formatter: function(params){
                                if (params.value > 0) {
                                    return params.value;
                                } else {
                                    return '';
                                }
                            }
                        }
                    },
//                    markPoint : {
//                         data : [
//                            {
//                                name: 'max',
//                                type: 'max'
//                            }
//                        ],
//                        animationThreshold :50
//                    },
                    markLine : {
                        data : [
                            {name: 'Target',yAxis: 50}
                        ]
                    },
                    data:var_LT_0400
                },   
                {
                    name:'0481 LT',
                    type:'bar',
                    label: {
                        normal: {
                            show: true,
                            position: 'top',
                            formatter: function(params){
                                if (params.value > 0) {
                                    return params.value;
                                } else {
                                    return '';
                                }
                            }
                        }
                    },
//                    markPoint : {
//                         data : [
//                            {
//                                name: 'max',
//                                type: 'max'
//                            }
//                        ],
//                        animationThreshold :50
//                    },
                    markLine : {
                        data : [
                            {name: 'Target',yAxis: 50}
                        ]
                    },
                    data:var_LT_0481
                },
                {
                    name:'0400+0481 LT',
                    type:'bar',
                    label: {
                        normal: {
                            show: true,
                            position: 'top',
                            formatter: function(params){
                                if (params.value > 0) {
                                    return params.value;
                                } else {
                                    return '';
                                }
                            }
                        }
                    },
//                    markPoint : {
//                         data : [
//                            {
//                                name: 'max',
//                                type: 'max'
//                            }
//                        ],
//                        animationThreshold :50
//                    },
                    markLine : {
                        data : [
                            {name: 'Target',yAxis: 50}
                        ]
                    },
                    data:var_LT
                }             
            ]
        }
        myChart6.hideLoading();    //隐藏加载动画
        myChart6.setOption(option6);
        //*****************************************************************************************************************************
        //**                                                      myChart7                                                           **
        //*****************************************************************************************************************************
        // 基于准备好的dom，初始化echarts图表
        var myChart7 = echarts.init(document.getElementById('divECharts7'),'FestoColor_6');
        // 过渡---------------------
        myChart7.showLoading({
            text: '正在努力的读取数据中...',    //loading话术
        });

        var var_Echarts7=values_ECharts7.split(';');

        var var_Reminder_Monitor_LT = var_Echarts7[0].split('!');  
        var var_Failed_Monitor = var_Echarts7[1].split('!'); 
          
        var values_Reminder_Monitor_LT = []; 
        var values_Failed_Monitor = []; 
        
        values_Reminder_Monitor_LT.push({ "value": var_Reminder_Monitor_LT[0], "name": '物料计划提醒' });
        values_Reminder_Monitor_LT.push({ "value": var_Reminder_Monitor_LT[1], "name": '生产计划提醒' }); 
        values_Failed_Monitor.push({ "value": var_Failed_Monitor[0], "name": 'DC Failed' });
        values_Failed_Monitor.push({ "value": var_Failed_Monitor[1], "name": 'Request Failed' }); 

        // 图表使用-------------------
        var option7 = {
                        title: [{
                            text: 'Reminder Monitor',
                            subtext: 'According to Quotation LT',       
                            x: '25%',
                            y: '10',
                            textAlign: 'center'
                        }, { 
                            text: 'Failed Monitor',       
                            subtext: 'Failed Order',
                            x: '75%',
                            y: '10',
                            textAlign: 'center'
                        }],
                        tooltip: {
                            trigger: 'item',
                            formatter: "{a} <br/>{b} : {c} ({d}%)"
                        },
                        legend: {
                            bottom: 10,
                            left: 'center',
                            data:  ['物料计划提醒','生产计划提醒','DC Failed','Request Failed']
                        },
                        series: [
                            {                                
                                name: '数量',
                                clockWise:false,
                                type: 'pie',
                                radius: '65%',
                                center: ['25%', '50%'],
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
                                data: values_Reminder_Monitor_LT
                            },  {                                
                                name: '数量',
                                clockWise:false,
                                type: 'pie',
                                radius: '65%',
                                center: ['75%', '50%'],
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
                                data: values_Failed_Monitor
                            },                        
                            ]
                    };         
       
        myChart7.hideLoading();    //隐藏加载动画
        myChart7.setOption(option7);
        //*****************************************************************************************************************************
        //**                                                      myChart8                                                           **
        //*****************************************************************************************************************************
        // 基于准备好的dom，初始化echarts图表
        var myChart8 = echarts.init(document.getElementById('divECharts8'),'FestoColor_6');
        // 过渡---------------------
        myChart8.showLoading({
            text: '正在努力的读取数据中...',    //loading话术
        });

        var var_Echarts8=values_ECharts8.split('!');
          
        var values_Overstocked_Monitor = []; 
        
        values_Overstocked_Monitor.push({ "value": var_Echarts8[0], "name": '<=5' });
        values_Overstocked_Monitor.push({ "value": var_Echarts8[1], "name": '5-10' }); 
        values_Overstocked_Monitor.push({ "value": var_Echarts8[2], "name": '10-20' });
        values_Overstocked_Monitor.push({ "value": var_Echarts8[3], "name": '>20' }); 
        values_Overstocked_Monitor.push({ "value": var_Echarts8[4], "name": '未发货' });

        // 图表使用-------------------
        var option8 = {
                        title: {
                            text: 'Overstocked Monitor',
                            x: 'center'
                        },
                        tooltip: {
                            trigger: 'item',
                            formatter: "{a} <br/>{b} : {c} ({d}%)"
                        },
                        legend: {
                            orient: 'vertical',
                            x: 'left',
                            data:  ['<=5','5-10','10-20','>20','未发货']
                        },
                        series: [
                            {                                
                                name: '数量',
                                clockWise:false,
                                type: 'pie',
                                radius: '65%',
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
                                        length1: 30
                                    }
                                },                    
                                data: values_Overstocked_Monitor
                            },                          
                            ]
                    };         
       
        myChart8.hideLoading();    //隐藏加载动画
        myChart8.setOption(option8);
        //*****************************************************************************************************************************
         /*窗口自适应，关键代码*/
        //*****************************************************************************************************************************
	    setTimeout(function (){
	        window.onresize = function () {
                myChart1.resize();	
                myChart2.resize();	
                myChart3.resize();   
                myChart4.resize();	
                myChart5.resize();	
                myChart6.resize(); 
                myChart7.resize(); 
                myChart8.resize();       	
	        }
	    },200)
        return false;
       }

    </script>
    </form>
</body>
</html>
