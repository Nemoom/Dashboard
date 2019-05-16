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
                    GCN CS Transparency kit
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
            <asp:Panel ID="Panel1" runat="server" Height="900">
                <table style="width: 100%;">
                    <tr style="height: 50%;">
                        <td style="width: 50%;">
                            <!-- 为ECharts准备一个具备大小（宽高）的Dom -->
                            <div id="divECharts1_1" style="height: 425px; border: 1px solid #ccc; padding: 10px; background-color: #E0E0E0;">
                            </div>
                        </td>
                        <td style="width: 50%;">
                            <!-- 为ECharts准备一个具备大小（宽高）的Dom -->
                            <div id="divECharts1_2" style="height: 425px; border: 1px solid #ccc; padding: 10px; background-color: #E0E0E0;">
                            </div>
                        </td>
                    </tr>
                    <tr style="height: 50%;">
                        <td style="width: 50%;">
                            <!-- 为ECharts准备一个具备大小（宽高）的Dom -->
                            <div id="divECharts1_3" style="height: 425px; border: 1px solid #ccc; padding: 10px;">
                            </div>
                        </td>
                        <td style="width: 50%;">
                            <!-- 为ECharts准备一个具备大小（宽高）的Dom -->
                            <div id="divECharts1_4" style="height: 425px; border: 1px solid #ccc; padding: 10px;">
                            </div>
                        </td>
                    </tr>
                </table>
            </asp:Panel>
            <asp:Panel ID="Panel2" runat="server" Height="900">
                <table style="width: 100%;">
                    <tr style="height: 50%;">
                        <td style="width: 50%;">
                            <!-- 为ECharts准备一个具备大小（宽高）的Dom -->
                            <div id="divECharts2_1" style="background-position: center center; height: 425px; border: 1px solid #ccc; 
                                padding: 10px; background-size: 80% 70%; background-image: url('./img/in.png'); background-repeat: no-repeat;">
                            <%--<div id="divECharts2_1" style="height: 500px; border: 1px solid #ccc; padding: 10px;">--%>
                            </div>
                        </td>
                        <td style="width: 50%;">
                            <!-- 为ECharts准备一个具备大小（宽高）的Dom -->
                            <div id="divECharts2_2" style="background-position: center center; height: 425px; border: 1px solid #ccc; 
                                padding: 10px; background-size: 80% 70%; background-image: url('./img/out.png'); background-repeat: no-repeat;">
                            </div>
                        </td>
                    </tr>
                    <tr style="height: 50%;">
                        <td style="width: 50%;">
                            <!-- 为ECharts准备一个具备大小（宽高）的Dom -->
                            <div id="divECharts2_3" style="height: 425px; border: 1px solid #ccc; padding: 10px;">
                            </div>
                        </td>
                        <td style="width: 50%;">
                            <!-- 为ECharts准备一个具备大小（宽高）的Dom -->
                            <div id="divECharts2_4" style="height: 425px; border: 1px solid #ccc; padding: 10px;">
                            </div>
                        </td>
                    </tr>
                </table>
            </asp:Panel>   
            <asp:Panel ID="Panel3" runat="server" Height="900">
                <table style="width: 100%;">
                    <tr style="height: 50%;">
                        <td style="width: 50%;">
                            <!-- 为ECharts准备一个具备大小（宽高）的Dom -->
                            <div id="divECharts3_1" style="height: 425px; border: 1px solid #ccc; padding: 10px;">
                            </div>
                        </td>
                        <td style="width: 50%;">
                            <!-- 为ECharts准备一个具备大小（宽高）的Dom -->
                            <div id="divECharts3_2" style="height: 425px; border: 1px solid #ccc; padding: 10px;">
                            </div>
                        </td>
                    </tr>
                    <tr style="height: 50%;">
                        <td style="width: 50%;">
                            <!-- 为ECharts准备一个具备大小（宽高）的Dom -->
                            <div id="divECharts3_3" style="height: 425px; border: 1px solid #ccc; padding: 10px;">
                            </div>
                        </td>
                        <td style="width: 50%;">
                            <!-- 为ECharts准备一个具备大小（宽高）的Dom -->
                            <div id="divECharts3_4" style="height: 425px; border: 1px solid #ccc; padding: 10px;">
                            </div>
                        </td>
                    </tr>
                </table>
            </asp:Panel>
            <asp:Label ID="Label1" runat="server" Text=""></asp:Label>            
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
//        document.getElementById('divECharts1_1').style.height=window.screen.width/4;
//        var x = document.getElementById('divECharts1_1').style.height;
//        var y = window.screen.width/4;
//        var msg = "divECharts1_1分辨率：" + x + "x" + y + "像素";
//        alert(msg);
//        document.getElementById('divECharts1_2').style.height=window.screen.width/4;
//        document.getElementById('divECharts1_3').style.height=window.screen.width/4;
//        document.getElementById('divECharts1_4').style.height=window.screen.width/4;
//        document.getElementById('divECharts2_1').style.height=window.screen.width/4;
//        document.getElementById('divECharts2_2').style.height=window.screen.width/4;
//        document.getElementById('divECharts2_3').style.height=window.screen.width/4;
//        document.getElementById('divECharts2_4').style.height=window.screen.width/4;
//        document.getElementById('divECharts3_1').style.height=window.screen.width/4;
//        document.getElementById('divECharts3_2').style.height=window.screen.width/4;
//        document.getElementById('divECharts3_3').style.height=window.screen.width/4;
//        document.getElementById('divECharts3_4').style.height=window.screen.width/4;
        // 基于准备好的dom，初始化echarts图表
        var myChart1_1 = echarts.init(document.getElementById('divECharts1_1'),'FestoColor_6');
        var myChart1_2 = echarts.init(document.getElementById('divECharts1_2'),'FestoColor_6');
        var myChart1_3 = echarts.init(document.getElementById('divECharts1_3'),'FestoColor_6');
        var myChart1_4 = echarts.init(document.getElementById('divECharts1_4'),'FestoColor_6');
        var myChart2_1 = echarts.init(document.getElementById('divECharts2_1'),'FestoColor_6');
        var myChart2_2 = echarts.init(document.getElementById('divECharts2_2'),'FestoColor_6');
        var myChart2_3 = echarts.init(document.getElementById('divECharts2_3'),'FestoColor_6');
        var myChart2_4 = echarts.init(document.getElementById('divECharts2_4'),'FestoColor_6');
        var myChart3_1 = echarts.init(document.getElementById('divECharts3_1'),'FestoColor_6');
        var myChart3_2 = echarts.init(document.getElementById('divECharts3_2'),'FestoColor_6');
        var myChart3_3 = echarts.init(document.getElementById('divECharts3_3'),'FestoColor_6');
        var myChart3_4 = echarts.init(document.getElementById('divECharts3_4'),'FestoColor_6');
        
        // 过渡---------------------
        myChart1_1.showLoading({
            text: '正在努力的读取数据中...',   
        });
        myChart1_2.showLoading({
            text: '正在努力的读取数据中...',   
        });        
        myChart1_3.showLoading({
            text: '正在努力的读取数据中...',  
        });
        myChart1_4.showLoading({
            text: '正在努力的读取数据中...',   
        });
        myChart2_1.showLoading({
            text: '正在努力的读取数据中...',   
        });
        myChart2_2.showLoading({
            text: '正在努力的读取数据中...',   
        });        
        myChart2_3.showLoading({
            text: '正在努力的读取数据中...',  
        });
        myChart2_4.showLoading({
            text: '正在努力的读取数据中...',   
        });
        myChart3_1.showLoading({
            text: '正在努力的读取数据中...',   
        });
        myChart3_2.showLoading({
            text: '正在努力的读取数据中...',   
        });        
        myChart3_3.showLoading({
            text: '正在努力的读取数据中...',  
        });
        myChart3_4.showLoading({
            text: '正在努力的读取数据中...',   
        });
    
        var count = 0;
        var time_count;
        time_count = setInterval(function () {
                count++;
                if (count%3==0) {                    
                    document.getElementById('Panel1').style.display="block";
                    document.getElementById('Panel2').style.display="none";
                    document.getElementById('Panel3').style.display="none";
                } 
                else if (count%3==1){
                    document.getElementById('Panel1').style.display="none";
                    document.getElementById('Panel2').style.display="block";
                    document.getElementById('Panel3').style.display="none";
                }   
                else {
                    document.getElementById('Panel1').style.display="none";
                    document.getElementById('Panel2').style.display="none";
                    document.getElementById('Panel3').style.display="block";
                }                  
            }, 20000);
    </script>
    <script type="text/javascript">
        function setECharts2(values_ECharts1_1,values_ECharts1_2,values_ECharts1_3,values_ECharts1_4,values_ECharts2_1,values_ECharts2_2,values_ECharts2_3,values_ECharts2_4,values_ECharts3_1,values_ECharts3_2,values_ECharts3_3,values_ECharts3_4) {    
        
        //*****************************************************************************************************************************
        //**                                                      myChart1_1                                                         **
        //*****************************************************************************************************************************
        // 基于准备好的dom，初始化echarts图表
        var myChart1_1 = echarts.init(document.getElementById('divECharts1_1'),'FestoColor_6');
        // 过渡---------------------
        myChart1_1.showLoading({
            text: '正在努力的读取数据中...',    //loading话术
        });

        var var_Echarts1_1=values_ECharts1_1.split(';');

        var var_DCR = var_Echarts1_1[0].split('!');  
        var var_DCR_0400 = var_Echarts1_1[1].split('!'); 
        var var_DCR_0481 = var_Echarts1_1[2].split('!');
          
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
        var option1_1 = {
            title: {
                        text: 'DCR(Delivery Class Reliablity)', 
//                        subtext: 'According to Quotation LT',
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
                data:['0400 DCR','0481 DCR','0400+0481 DCR'],
                selected:{'0400 DCR': false ,'0481 DCR': false}
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
//                    itemStyle:{
//                        normal: {
//                            //每根柱子颜色设置
//                            borderColor: "blue"
//                        }
//                    },
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
//                    itemStyle:{
//                        normal: {
//                            //每根柱子颜色设置
//                            color: function(params) {
//                                var colorList = [
//                                    "#b6c0c6",
//                                    "#b6c0c6",
//                                    "#b6c0c6",
//                                    "#b6c0c6",
//                                    "#b6c0c6",
//                                    "#b6c0c6",
//                                    "#b6c0c6",
//                                    "#b6c0c6",
//                                    "#b6c0c6",
//                                    "#b6c0c6",
//                                    "#b6c0c6",
//                                    "#b6c0c6",                                    
//                                    "gray"
//                                ];
//                                return colorList[params.dataIndex];
//                            }
//                        }
//                    },
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
        myChart1_1.hideLoading();    //隐藏加载动画
        myChart1_1.setOption(option1_1);
        //*****************************************************************************************************************************
        //**                                                      myChart1_2                                                         **
        //*****************************************************************************************************************************
         // 基于准备好的dom，初始化echarts图表
        var myChart1_2 = echarts.init(document.getElementById('divECharts1_2'),'FestoColor_6');
        // 过渡---------------------
        myChart1_2.showLoading({
            text: '正在努力的读取数据中...',    //loading话术
        });
     
        var var_Echarts1_2=values_ECharts1_2.split(';');

        var var_LT = var_Echarts1_2[0].split('!');  
        var var_LT_0400 = var_Echarts1_2[1].split('!'); 
        var var_LT_0481 = var_Echarts1_2[2].split('!');
        var option1_2 = {
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
        myChart1_2.hideLoading();    //隐藏加载动画
        myChart1_2.setOption(option1_2);
        //*****************************************************************************************************************************
        //**                                                      myChart1_3                                                         **
        //*****************************************************************************************************************************
        // 基于准备好的dom，初始化echarts图表
        var myChart1_3 = echarts.init(document.getElementById('divECharts1_3'),'FestoColor_6');
        // 过渡---------------------
        myChart1_3.showLoading({
            text: '正在努力的读取数据中...',    //loading话术
        });

        var var_Echarts1_3=values_ECharts1_3.split(';');

        var var_Reminder_Monitor_LT = var_Echarts1_3[0].split('!');  
        var var_Failed_Monitor = var_Echarts1_3[1].split('!'); 
          
        var values_Reminder_Monitor_LT = []; 
        var values_Failed_Monitor = []; 
        
        values_Reminder_Monitor_LT.push({ "value": var_Reminder_Monitor_LT[0], "name": 'Material Planning Alert' });
        values_Reminder_Monitor_LT.push({ "value": var_Reminder_Monitor_LT[1], "name": 'Production Planning Alert' }); 
        values_Failed_Monitor.push({ "value": var_Failed_Monitor[0], "name": 'DC Failed & Match Req' });
        values_Failed_Monitor.push({ "value": var_Failed_Monitor[1], "name": 'DC Failed & Req Failed' }); 

        // 图表使用-------------------
        var option1_3 = {
                        title: [{
                              text: 'DCR ALERT',
                              subtext:'0400 当月',
                              x: 'center'
//                            text: 'Reminder Monitor',
//                            subtext: 'According to Quotation LT',       
//                            x: '25%',
//                            y: '10',
//                            textAlign: 'center'
//                        }, { 
//                            text: 'Failed Monitor',       
//                            subtext: 'Failed Order',
//                            x: '75%',
//                            y: '10',
//                            textAlign: 'center'
                        }],
                        tooltip: {
//                            trigger: 'item',
                            formatter: "{a} <br/>{b} : {c} ({d}%)"
                        },
                        toolbox: {
                            feature: {
                                dataView: {show: true, readOnly: false},
                                magicType: {show: false},
                                restore: {show: true},
                                saveAsImage: {show: true}
                            }
                        },
                        legend: {
                            bottom: 10,
                            left: 'center',
                            data:  ['Material Planning Alert','Production Planning Alert','DC Failed & Match Req','DC Failed & Req Failed']
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
                                        formatter: '{b|{b}：{c}} \n {per|{d}%}  ',                                        
                                        position:'inside',
                                        rich: {                                            
                                            b: {
                                                color: '#eee',
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
                                        formatter: '{b|{b}：{c}} \n {per|{d}%}  ',                                        
                                        position:'inside',
                                        rich: {                                            
                                            b: {
                                                color: '#eee',
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
                                        show: false
                                    }
                                },                
                                data: values_Failed_Monitor
                            },                        
                            ]
                    };         
       
        myChart1_3.hideLoading();    //隐藏加载动画
        myChart1_3.setOption(option1_3);        
        //*****************************************************************************************************************************
        //**                                                      myChart1_4                                                         **
        //*****************************************************************************************************************************
        // 基于准备好的dom，初始化echarts图表
        var myChart1_4 = echarts.init(document.getElementById('divECharts1_4'),'FestoColor_6');
        // 过渡---------------------
        myChart1_4.showLoading({
            text: '正在努力的读取数据中...',    //loading话术
        });

        var var_Echarts1_4=values_ECharts1_4.split('!');
          
        var values_Ongoing_Monitor = []; 
        
        values_Ongoing_Monitor.push({ "value": var_Echarts1_4[0], "name": '<=40 CDS' });
        values_Ongoing_Monitor.push({ "value": var_Echarts1_4[1], "name": '40-50 CDS' }); 
        values_Ongoing_Monitor.push({ "value": var_Echarts1_4[2], "name": '>50 CDS' });

        // 图表使用-------------------
        var option1_4 = {
                        title: {
                            text: 'Ongoing Pro. Order Monitoring',
                            subtext: '0400 YTD',    
                            x: 'center'
                        },
                        tooltip: {
                            trigger: 'item',
                            formatter: "{a} <br/>{b} : {c} ({d}%)"
                        },
                        legend: {
                            orient: 'vertical',
                            x: 'left',
                            data:  ['<=40 CDS','40-50 CDS','>50 CDS']
                        },
                        toolbox: {
                            feature: {
                                dataView: {show: true, readOnly: false},
                                magicType: {show: false},
                                restore: {show: true},
                                saveAsImage: {show: true}
                            }
                        },
                        series: [
                            {                                
                                name: '数量',
                                clockWise:false,
                                type: 'pie',
                                radius: '65%',
                                center: ['50%', '50%'],
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
                                data: values_Ongoing_Monitor
                            },                          
                            ]
                    };         
       
        myChart1_4.hideLoading();    //隐藏加载动画
        myChart1_4.setOption(option1_4);
        //*****************************************************************************************************************************
        //**                                                      myChart2_1                                                         **
        //*****************************************************************************************************************************
         // 基于准备好的dom，初始化echarts图表
        var myChart2_1 = echarts.init(document.getElementById('divECharts2_1'),'FestoColor_6');
        // 过渡---------------------
        myChart2_1.showLoading({
            text: '正在努力的读取数据中...',    //loading话术
        });

        var var_Echarts2_1=values_ECharts2_1.split(';');

        var count_SO_soCreated_0400_perMonth=var_Echarts2_1[0].split('!');
        var count_SO_soCreated_0481_perMonth = var_Echarts2_1[1].split('!'); 
        var NetValue_SO_soCreated_perMonth = var_Echarts2_1[2].split('!'); 

        var count_SO_soCreated_0400_perMonth1=count_SO_soCreated_0400_perMonth.splice(0,count_SO_soCreated_0400_perMonth.length-1);
        var count_SO_soCreated_0481_perMonth1=count_SO_soCreated_0481_perMonth.splice(0,count_SO_soCreated_0481_perMonth.length-1);
        var NetValue_SO_soCreated_perMonth1 = NetValue_SO_soCreated_perMonth.splice(0,NetValue_SO_soCreated_perMonth.length-1);  
       
        var option2_1 = {
            title: {
                        text: 'Monthly SO Input', 
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
                data:['0400新进订单量','0481新进订单量','订单金额'],
                selected: {
                '订单金额': false
                }
            },   
//            graphic: [
//                {
//                    type: 'image',
//                    id: 'logo',
//                    right: 20,
//                    top: 20,
////                    z: -10,
//                    bounding: 'raw',
//                    origin: [75, 75],
//                    style: {
//                        image: 'image:///WebSite_DataBoard/img/FestoLogo.png',
//                        width: 150,
//                        height: 150
////                        opacity: 0.4
//                    }
//                }
            graphic: [
//                {
//                    type: 'image',
//                    id: 'logo',
//                    right: 20,
//                    top: 20,
//                    z: -10,
//                    bounding: 'raw',
//                    origin: [75, 75],
//                    style: {
//                        image: 'http://echarts.baidu.com/images/favicon.png',
//                        width: 150,
//                        height: 150,
//                        opacity: 0.4
//                    }
//                },
//                {
//                    type: 'group',
//                    rotation: Math.PI / 4,
//                    bounding: 'raw',
//                    right: 110,
//                    bottom: 110,
//                    z: 100,
//                    children: [
//                        {
//                            type: 'rect',
//                            left: 'center',
//                            top: 'center',
//                            z: 100,
//                            shape: {
//                                width: 400,
//                                height: 50
//                            },
//                            style: {
//                                fill: 'rgba(0,0,0,0.3)'
//                            }
//                        },
//                        {
//                            type: 'text',
//                            left: 'center',
//                            top: 'center',
//                            z: 100,
//                            style: {
//                                fill: '#fff',
//                                text: 'Sale Order Input',
//                                font: 'bold 26px Microsoft YaHei'
//                            }
//                        }
//                    ]
//                }
            ],         
            xAxis: [
                {
                    splitLine:{show: false},//去除网格线
                    type: 'category',
                    data: ['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月'],
                    axisPointer: {
                        type: 'shadow'
                    }
                }
            ],
            yAxis: [
                {
                    splitLine:{show: false},//去除网格线
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
                    visiable: false,
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
                    name:'0400新进订单量',
                    stack: '新进订单量',
                    type:'bar',
                    label: {
                        normal: {
                            show: true,
                            position: 'insideTop'
                        }
                    },
                    data:count_SO_soCreated_0400_perMonth1
                },   
                {
                    name:'0481新进订单量',
                    stack: '新进订单量',
                    type:'bar',
                    label: {
                        normal: {
                            show: true,
                            position: 'insideTop'
                        }
                    },
                    data:count_SO_soCreated_0481_perMonth1
                },              
                {
                    name:'订单金额',
                    type:'line',
                    yAxisIndex: 1,
                    data:NetValue_SO_soCreated_perMonth1
        
                },
                {
                    name:'合计',
                    type:'bar',
                    stack: '新进订单量',
                    label: {
                        normal: {
                            show: true,
                            position: 'top',
                            textStyle: {
//                                color: '#000'
                            },
                            formatter:''
                        }
                    },
                    data:[0,0,0,0,0,0,0,0,0,0,0,0]
                }
            ]
        }
        var series2_1 = option2_1["series"];
        var fun2_1 = function (params) { 
                var data2_1 =0;
                for(var i=0,l=2;i<l;i++){ 
                    data2_1 += parseInt(series2_1[i].data[params.dataIndex])
                } 
                if (data2_1 > 0) {
                    return data2_1;
                } else {
                    return '';
                }
            }
        //加载页面时候替换最后一个series的formatter
        series2_1[series2_1.length-1]["label"]["normal"]["formatter"] = fun2_1
        myChart2_1.hideLoading();    //隐藏加载动画
        myChart2_1.setOption(option2_1);  
        //*****************************************************************************************************************************
        //**                                                      myChart2_2                                                         **
        //*****************************************************************************************************************************
         // 基于准备好的dom，初始化echarts图表
        var myChart2_2 = echarts.init(document.getElementById('divECharts2_2'),'FestoColor_6');
        // 过渡---------------------
        myChart2_2.showLoading({
            text: '正在努力的读取数据中...',    //loading话术
        });

        var var_Echarts2_2=values_ECharts2_2.split(';');

        var count_PO_Finished_0400_perMonth=var_Echarts2_2[0].split('!');
        var count_PO_Finished_0481_perMonth = var_Echarts2_2[1].split('!');   
        var count_PO_ForecastFinished_perMonth = var_Echarts2_2[2].split('!');  
        var NetValue_PO_Finished_perMonth = var_Echarts2_2[3].split('!');   
        var count_PO_Finished_perMonth = var_Echarts2_2[4].split('!');   

        var count_PO_Finished_0400_perMonth1=count_PO_Finished_0400_perMonth.splice(0,count_PO_Finished_0400_perMonth.length-1);
        var count_PO_Finished_0481_perMonth1=count_PO_Finished_0481_perMonth.splice(0,count_PO_Finished_0481_perMonth.length-1);
        var count_PO_ForecastFinished_perMonth1=count_PO_ForecastFinished_perMonth.splice(0,count_PO_ForecastFinished_perMonth.length-1);
        var NetValue_PO_Finished_perMonth1 = NetValue_PO_Finished_perMonth.splice(0,NetValue_PO_Finished_perMonth.length-1);
        var count_PO_Finished_perMonth1=count_PO_Finished_perMonth.splice(0,count_PO_Finished_perMonth.length-1);

        var option2_2 = {
            title: {
                        text: 'Monthly Finished Pro. Order', 
                        x: 'center'
                    },
            tooltip: {
//                trigger: 'axis',
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
                '预计完成量': false,'订单金额':false
                },
                data:['0400订单完成量','0481订单完成量','预计完成量','订单金额']
            },
//            graphic: [
//                {
//                    type: 'group',
//                    rotation: Math.PI / 4,
//                    bounding: 'raw',
//                    right: 110,
//                    bottom: 110,
//                    z: 100,
//                    children: [
//                        {
//                            type: 'rect',
//                            left: 'center',
//                            top: 'center',
//                            z: 100,
//                            shape: {
//                                width: 400,
//                                height: 50
//                            },
//                            style: {
//                                fill: 'rgba(0,0,0,0.3)'
//                            }
//                        },
//                        {
//                            type: 'text',
//                            left: 'center',
//                            top: 'center',
//                            z: 100,
//                            style: {
//                                fill: '#fff',
//                                text: 'Finished Pro. Order',
//                                font: 'bold 26px Microsoft YaHei'
//                            }
//                        }
//                    ]
//                }
//            ], 
            xAxis: [
                {
                    type: 'category',
                    splitLine:{show: false},//去除网格线
                    data: ['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月'],
                    axisPointer: {
                        type: 'shadow'
                    }
                }
            ],
            yAxis: [
                {
                    type: 'value',
                    splitLine:{show: false},//去除网格线
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
                    visiable: false,
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
                    name:'0400订单完成量',
                    type:'bar',
                    stack: '完成量',
                    label: {
                        normal: {
                            show: true,
                            position: 'insideTop'
                        }
                    },
                    data:count_PO_Finished_0400_perMonth1
                },
                {
                    name:'0481订单完成量',
                    type:'bar',
                    stack: '完成量',
                    label: {
                        normal: {
                            show: true,
                            position: 'insideTop'
                        }
                    },
                    data:count_PO_Finished_0481_perMonth1
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
        
                },
                {
                    name:'合计',
                    type:'bar',
                    stack: '完成量',
                    label: {
                        normal: {
                            show: true,
                            position: 'top',
                            textStyle: {
//                                color: '#000'
                            },
                            formatter:''
                        }
                    },
                    data:[0,0,0,0,0,0,0,0,0,0,0,0]
                }
            ]
        }
        var series = option2_2["series"];
        var fun = function (params) { 
                var data3 =0;
                for(var i=0,l=2;i<l;i++){ 
                    data3 += parseInt(series[i].data[params.dataIndex])
                } 
                if (data3 > 0) {
                    return data3;
                } else {
                    return '';
                }
                
            }
        //加载页面时候替换最后一个series的formatter
        series[series.length-1]["label"]["normal"]["formatter"] = fun
        myChart2_2.hideLoading();    //隐藏加载动画
        myChart2_2.setOption(option2_2);      
        //*****************************************************************************************************************************
        //**                                                      myChart2_3                                                         **
        //*****************************************************************************************************************************
        // 基于准备好的dom，初始化echarts图表
        var myChart2_3 = echarts.init(document.getElementById('divECharts2_3'),'FestoColor_6');
        // 过渡---------------------
        myChart2_3.showLoading({
            text: '正在努力的读取数据中...',    //loading话术
        });

        var var_Echarts2_3=values_ECharts2_3.split('!');
          
        var values_Overstocked_Monitor = []; 
        
        values_Overstocked_Monitor.push({ "value": var_Echarts2_3[0], "name": '<=5 WD' });
        values_Overstocked_Monitor.push({ "value": var_Echarts2_3[1], "name": '5-10 WD' }); 
        values_Overstocked_Monitor.push({ "value": var_Echarts2_3[2], "name": '10-20 WD' });
        values_Overstocked_Monitor.push({ "value": var_Echarts2_3[3], "name": '>20 WD' }); 
        values_Overstocked_Monitor.push({ "value": var_Echarts2_3[4], "name": 'Undelivered' });

        // 图表使用-------------------
        var option2_3 = {
                        title: {
                            text: 'CS Shipment Statistic',
                            subtext:'0400+0481 YTD',
                            x: 'center'
                        },
                        tooltip: {
                            trigger: 'item',
                            formatter: "{a} <br/>{b} : {c} ({d}%)"
                        },
                        legend: {
                            orient: 'vertical',
                            x: 'left',
                            data:  ['<=5 WD','5-10 WD','10-20 WD','>20 WD','Undelivered']
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
       
        myChart2_3.hideLoading();    //隐藏加载动画
        myChart2_3.setOption(option2_3);
        //*****************************************************************************************************************************
        //**                                                      myChart2_4                                                         **
        //*****************************************************************************************************************************
        // 基于准备好的dom，初始化echarts图表
        var myChart2_4 = echarts.init(document.getElementById('divECharts2_4'),'FestoColor_6');
        // 过渡---------------------
        myChart2_4.showLoading({
            text: '正在努力的读取数据中...',    //loading话术
        });

        var var_Echarts2_4=values_ECharts2_4.split('!');
          
        var values_Ready4Shipment_Monitor = []; 
        
        values_Ready4Shipment_Monitor.push({ "value": var_Echarts2_4[0], "name": '<=10 WD' });
        values_Ready4Shipment_Monitor.push({ "value": var_Echarts2_4[1], "name": '10-20 WD' }); 
        values_Ready4Shipment_Monitor.push({ "value": var_Echarts2_4[2], "name": '20-30 WD' });
        values_Ready4Shipment_Monitor.push({ "value": var_Echarts2_4[3], "name": '>30 WD' }); 

        // 图表使用-------------------
        var option2_4 = {
                        title: {
                            text: 'Undelivered Monitoring',
                            subtext:'0400+0481',
                            x: 'center'
                        },
                        tooltip: {
                            trigger: 'item',
                            formatter: "{a} <br/>{b} : {c} ({d}%)"
                        },
                        legend: {
                            orient: 'vertical',
                            x: 'left',
                            data:  ['<=10 WD','10-20 WD','20-30 WD','>30 WD']
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
                                data: values_Ready4Shipment_Monitor
                            },                          
                            ]
                    };         
       
        myChart2_4.hideLoading();    //隐藏加载动画
        myChart2_4.setOption(option2_4);
        //*****************************************************************************************************************************
        //**                                                      myChart3_1                                                         **
        //*****************************************************************************************************************************
        // 基于准备好的dom，初始化echarts图表
        var myChart3_1 = echarts.init(document.getElementById('divECharts3_1'),'FestoColor_6');
        // 过渡---------------------
        myChart3_1.showLoading({
            text: '正在努力的读取数据中...',    //loading话术
        });

        var var_Echarts3_1=values_ECharts3_1.split(';');

        var count_PO_soCreated_perMonth = var_Echarts3_1[0].split('!');  
        var count_CreatedDelay_perMonth = var_Echarts3_1[1].split('!'); 
        var PO_CreationMonitor_YTD = var_Echarts3_1[2].split('!');   
        var values_PO_CreationMonitor_YTD = [];    
        values_PO_CreationMonitor_YTD.push({ "value": PO_CreationMonitor_YTD[0], "name": '正常创建PO数(YTD)' });
        values_PO_CreationMonitor_YTD.push({ "value": PO_CreationMonitor_YTD[1], "name": '延期创建PO数(YTD)' });        

        // 图表使用-------------------
        var option3_1 = {
                        title: {
                            text: '0400 SO → PO Creation',  
                            subtext: 'Target：LT < 1(WD)',
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
                            data:['总创建PO数','其中延期创建PO数']
                            
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
                                name:'总创建PO数',
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
                                data:count_PO_soCreated_perMonth
                            },
                            {
                                name:'其中延期创建PO数',
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
                                label: {
                                    normal: {
                                        formatter: '{a|  {b}  }{abg|}\n{hr|}\n  {c}  {per|{d}%}  ',
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
                                data:values_PO_CreationMonitor_YTD                                
                            }
                        ]
                    };         
       
        myChart3_1.hideLoading();    //隐藏加载动画
        myChart3_1.setOption(option3_1);
        //*****************************************************************************************************************************
        //**                                                      myChart3_2                                                         **
        //*****************************************************************************************************************************
        // 基于准备好的dom，初始化echarts图表
        var myChart3_2 = echarts.init(document.getElementById('divECharts3_2'),'FestoColor_6');
        // 过渡---------------------
        myChart3_2.showLoading({
            text: '正在努力的读取数据中...',    //loading话术
        });

        var var_Echarts3_2=values_ECharts3_2.split(';');

        var count_PO_poCreated_perMonth = var_Echarts3_2[0].split('!');  
        var count_ReleaseDelay_perMonth = var_Echarts3_2[1].split('!'); 
        var PO_Release_YTD = var_Echarts3_2[2].split('!');   
        var values_PO_Release_YTD = [];    
        values_PO_Release_YTD.push({ "value": PO_Release_YTD[0], "name": '按时释放PO数(YTD)' });
        values_PO_Release_YTD.push({ "value": PO_Release_YTD[1], "name": '延期释放PO数(YTD)' });        

        // 图表使用-------------------
        var option3_2 = {
                        title: {
                            text: '0400 PO Creation → PO Release', 
                            subtext: 'Target：LT <= 5(WD)',
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
                            data:['总释放PO数','其中延期释放PO数']
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
                                name:'总释放PO数',
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
                                data:count_PO_poCreated_perMonth
                            },
                            {
                                name:'其中延期释放PO数',
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
                                label: {
                                    normal: {
                                        formatter: '{a|  {b}  }{abg|}\n{hr|}\n  {c}  {per|{d}%}  ',
                                        backgroundColor: '#eee',
                                        borderColor: '#aaa',
                                        borderWidth: 1,
                                        borderRadius: 4,
                                        rich: {
                                            a: {
                                                color: '#999',
//                                                fontSize: 16,
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
                                data:values_PO_Release_YTD                                
                            }
                        ]
                    };         
       
        myChart3_2.hideLoading();    //隐藏加载动画
        myChart3_2.setOption(option3_2);
        //*****************************************************************************************************************************
        //**                                                      myChart3_3                                                         **
        //*****************************************************************************************************************************
         // 基于准备好的dom，初始化echarts图表
        var myChart3_3 = echarts.init(document.getElementById('divECharts3_3'),'FestoColor_6');
        // 过渡---------------------
        myChart3_3.showLoading({
            text: '正在努力的读取数据中...',    //loading话术
        });
     
        var var_Echarts3_3=values_ECharts3_3.split(';');

        var var_POFinish = var_Echarts3_3[0].split('!');  
        var var_POFinish_0400 = var_Echarts3_3[1].split('!'); 
        var var_POFinish_0481 = var_Echarts3_3[2].split('!');
        var option3_3 = {
            title: {
                        text: '0400 PO Release → Actual finish', 
                        subtext:'Target：23WD',
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
                data:['0400','0481','0400+0481'],
                selected: {
                '0481': false, '0400+0481': false
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
                    name: 'Average Workday',
//                    min: 0,
//                    max: 1000,
//                    interval: 50,
                    axisLabel: {
                        formatter: '{value}WD'
                    }
                },               
            ],
            series: [                 
                {
                    name:'0400',
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
                            {name: 'Target',yAxis: 23}
                        ]
                    },
                    data:var_POFinish_0400
                },   
                {
                    name:'0481',
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
                            {name: 'Target',yAxis: 23}
                        ]
                    },
                    data:var_POFinish_0481
                },
                {
                    name:'0400+0481',
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
                            {name: 'Target',yAxis: 23}
                        ]
                    },
                    data:var_POFinish
                }             
            ]
        }
        myChart3_3.hideLoading();    //隐藏加载动画
        myChart3_3.setOption(option3_3);
        //*****************************************************************************************************************************
        //**                                                      myChart3_4                                                         **
        //*****************************************************************************************************************************
         // 基于准备好的dom，初始化echarts图表
        var myChart3_4 = echarts.init(document.getElementById('divECharts3_4'),'FestoColor_6');
        // 过渡---------------------
        myChart3_4.showLoading({
            text: '正在努力的读取数据中...',    //loading话术
        });
     
        var var_Echarts3_4=values_ECharts3_4.split(';');

        var var_Ex_plant = var_Echarts3_4[0].split('!');  
        var var_Ex_plant_0400 = var_Echarts3_4[1].split('!'); 
        var var_Ex_plant_0481 = var_Echarts3_4[2].split('!');
        var option3_4 = {
            title: {
                        text: '0400 Actual finish → Ex-plant', 
                        subtext:'Target：2WD',
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
                data:['0400','0481','0400+0481'],
                selected: {
                '0481': false, '0400+0481': false
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
                    name: 'Average Workday',
//                    min: 0,
//                    max: 1000,
//                    interval: 50,
                    axisLabel: {
                        formatter: '{value}WD'
                    }
                },               
            ],
            series: [                 
                {
                    name:'0400',
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
                            {name: 'Target',yAxis: 2}
                        ]
                    },
                    data:var_Ex_plant_0400
                },   
                {
                    name:'0481',
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
                            {name: 'Target',yAxis: 2}
                        ]
                    },
                    data:var_Ex_plant_0481
                },
                {
                    name:'0400+0481',
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
                            {name: 'Target',yAxis: 2}
                        ]
                    },
                    data:var_Ex_plant
                }             
            ]
        }
        myChart3_4.hideLoading();    //隐藏加载动画
        myChart3_4.setOption(option3_4);
        //*****************************************************************************************************************************
         /*窗口自适应，关键代码*/
        //*****************************************************************************************************************************
	    setTimeout(function (){
	        window.onresize = function () {
                myChart1_1.resize();	
                myChart1_2.resize();	
                myChart1_3.resize();
                myChart1_4.resize();   
                myChart2_1.resize();	
                myChart2_2.resize();	
                myChart2_3.resize(); 
                myChart2_4.resize();
                myChart3_1.resize(); 
                myChart3_2.resize();
                myChart3_3.resize(); 
                myChart3_4.resize();       	
	        }
	    },200)
        return false;
       }

    </script>
    </form>
</body>
</html>
