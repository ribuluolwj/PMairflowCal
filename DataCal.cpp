// Copyright 2008...
// License(BSD/GPL...)
// Author: RenPeng
// This is
#include <iostream>
#include <algorithm>
#include <cmath>
#include <map>
#include <conio.h>
//#include <windows.h>
#include "libxl.h"

// 设定TIP系数纸种类型和系数个数
#define paperCount 25
#define factorCount 30

using namespace std;
using namespace libxl;

// 格式化输出字符串，左对齐，参数：字符串，填充类型(默任“空格”），字符串长度（默认60）
string formatString(string str, string slot = " ", int length = 60)
{
    string sb;
    sb.append(str);
    int count = length - str.length();
    while(count > 0){
        sb.append(slot);
        count --;
    }
    return sb;
}
//中文的内容读出来后要进行编码的转换，这个为转换函数：wchar_t to char
//char *w2c(char *pcstr, const wchar_t *pwstr, size_t len)
//{
//    int nlength = wcslen(pwstr);
//    //获取转换后的长度
//    int nbytes = WideCharToMultiByte(0, 0, pwstr, nlength, NULL, 0, NULL, NULL);
//    if (nbytes>len)   nbytes = len;
    // 通过以上得到的结果，转换unicode 字符为ascii 字符
//    WideCharToMultiByte(0, 0, pwstr, nlength, pcstr, nbytes, NULL, NULL);
//    return pcstr;
//}

// 计算主程序
int main()
{
    // 控制台显示乱码纠正
    // system("chcp 65001"); //设置字符集，最终编译时加上，否则调试时出问题
    // 设定纸种系数字典
    std::map<string, int> paperID =
        {
            {"CORRUGATING MEDIUM",          0},
            {"PRINTING & WRITING",          1},
            {"BOOK PAPERS",                 2},
            {"HEAVYWEIGHTS",                3},
            {"MG & MF PAPERS",              4},
            {"GLASSINE, GREASEPROOF",       5},
            {"CARBONIZING",                 6},
            {"WAXING BASE",                 7},
            {"CIGARETTE, CONDENSOR TISSUE", 8},
            {"NAPKIN",                      9},
            {"TOWEL, TWO PLY",             10},
            {"TOWEL, SINGLE PLY",          11},
            {"DIRECTORY ROTO, CATALOG",    12},
            {"SC MAGAZINE",                13},
            {"LWC PUBLICATION",            14},
            {"NEWSPRINT",                  15},
            {"BAG",                        16},
            {"SATURATING",                 17},
            {"LINERBOARD-LW",              18},
            {"LINERBOARD-HW",              19},
            {"SBS PAPERBOARD",             20},
            {"PACKAGING SPECIALTIES",      21},
            {"PULP",                       22}};
    // 定义数据表中：真空元件信息起始列号、写入数据起始列号、纸机真空元件划分的段数
    int startColnum, writeColnum, elementNum;
    startColnum =  4; // 真空元件抽吸位置名称列开始
    writeColnum = 13; // M列开始
    elementNum  = 25; // unitLocation数值范围1~25，意思是纸机真空元件划分了25种不同类型
    // 定义计算公式区分位置的工段号
    int boxPart, flatboxPart, rollPartstart, rollPartstop, feltboxPart;
    boxPart       =  3; // 吸湿箱最大工段号，1~3全部使用粗算公式，并不考虑实际暴露面积，至考虑元件数量，真空度低，无需校准
    flatboxPart   =  5; // 真空箱最大工段号，4~5，分别为低真空和高真空吸水箱，4之后是水线，此时计算公式考虑暴露面积
    rollPartstart =  6; // 真空吸辊开始工段号，6~21，分别为伏辊低真空、高真空，吸移辊低、高真空，压榨辊低真空、高真空
    rollPartstop  = 21; // 真空吸辊结束工段号
    feltboxPart   = 22; // 毛毯吸水箱和转移箱开始工段号，22~25
    // 定义纸种类型对应的数组索引标号
    int vFi = 798;
    // 定义：需要计算的纸种类型
    string paperTypeset;

    // 定义纸种，确定计算数据
    cout << "[1 ] 瓦楞芯纸（CORRUGATING MEDIUM）；\n"
         << "[2 ] 复印纸和书写纸（PRINTING & WRITING）；\n"
         << "[3 ] 书本纸（BOOK PAPERS）；\n"
         << "[4 ] 高定量印刷纸（HEAVYWEIGHTS）；\n"
         << "[5 ] MG & MF（MG & MF PAPERS）；\n"
         << "[6 ] 格拉辛纸（GLASSINE, GREASEPROOF）；\n"
         << "[7 ] 复写原纸（CARBONIZING）；\n"
         << "[8 ] 蜡基特纸（WAXING BASE）；\n"
         << "[9 ] 卷烟纸（CIGARETTE, CONDENSOR TISSUE）；\n"
         << "[10] 餐巾纸（NAPKIN）；\n"
         << "[11] 双层毛巾纸（TOWEL, TWO PLY）；\n"
         << "[12] 单层毛巾纸（TOWEL, SINGLE PLY）；\n"
         << "[13] 目录纸（DIRECTORY ROTO, CATALOG）；\n"
         << "[14] SC杂志纸（SC MAGAZINE）；\n"
         << "[15] LWC出版纸（LWC PUBLICATION）；\n"
         << "[16] 新闻纸（NEWSPRINT）；\n"
         << "[17] 纸袋纸（BAG）；\n"
         << "[18] 饱和纸（SATURATING）；\n"
         << "[19] 低定量线板纸（LINERBOARD-LW）；\n"
         << "[20] 高定量线板纸（LINERBOARD-HW）；\n"
         << "[21] 白芯白卡纸（SBS PAPERBOARD）；\n"
         << "[22] 特种包装纸（PACKAGING SPECIALTIES）；\n"
         << "[23] 浆板（PULP）。" << endl;
    cout << "请输入计算纸种参数（若同时计算多种纸种，或原始数据表中已选定，可跳过输入）：";

    while (1)
    {
        // 定义临时字符数组用来存取屏幕读入的数据
        char pTtemp[6] = "798";
        cin.getline(pTtemp, 6);
        // 将读取的输入字符数组转换为字符串
        string pT = pTtemp;
        // 定义输入数据正确性判定字符串
        string ans = "1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23";
        string::size_type idx;
        // 查找输入数据是否在判定字符串中
        idx = ans.find(pT);
        // 判定输入是否正确
        if (idx == string::npos || pT.length() > 2)
        {
            cout << "输入有误，请重新输入：" << endl;
        }
        else
        {
            if (pT == "")
            {
                break;
            }
            else
            {
                vFi = stoi(pT) - 1;
                break;
            }
        }
    }
    int processControl = 3;
    cout << "[1] 临时计算，计算值只显示在屏幕上；\n"
         << "[2] 存储结果，计算值存储到原Excel；\n"
         << "[3] 存储公式，计算公式写入Excel。" << endl;
    cout << "请输入进程参数（默认值3）：";
    while (1)
    {
        // 定义临时字符数组用来存取屏幕读入的数据
        char pCtemp[6] = "3";
        cin.getline(pCtemp, 6);
        // 将读取的输入字符数组转换为字符串
        string pC = pCtemp;
        // 定义输入数据正确性判定字符串
        string ans = "3,2,1";
        string::size_type idx;
        // 查找输入数据是否在判定字符串中
        idx = ans.find(pC);
        // 判定输入是否正确
        if (idx == string::npos || pC.length() > 1)
        {
            cout << "输入有误，请重新输入：" << endl;
        }
        else
        {
            if (pC == "")
            {
                break;
            }
            else
            {
                processControl = stoi(pC);
                break;
            }
        }
    }


    // 暂时设定计算的纸种
    // paperTypeset = "CORRUGATING MEDIUM";
    // 根据纸种类型转换成数组索引标号
    // vFi = paperID.at(paperTypeset);

    // 分配二维动态数组：存储TIP真空抽气系数最小值
    double **vacuumMinFactor = NULL;
    vacuumMinFactor = new double *[paperCount]; // 开辟动态数组行
    for (int i = 0; i < paperCount; i++) // 开辟动态数组列
    {
        vacuumMinFactor[i] = new double [factorCount];
    }
    // 分配二维动态数组：存储TIP真空度参考值最小值
    double **referMinVacuum = NULL; // 开辟动态数组行
    referMinVacuum = new double *[paperCount]; // 开辟动态数组行
    for (int i = 0; i < paperCount; i++) // 开辟动态数组列
    {
        referMinVacuum[i] = new double[factorCount];
    }
    // 分配二维动态数组：存储TIP真空抽气系数最大值
    double **vacuumMaxFactor = NULL;
    vacuumMaxFactor = new double *[paperCount]; // 开辟动态数组行
    for (int i = 0; i < paperCount; i++) // 开辟动态数组列
    {
        vacuumMaxFactor[i] = new double [factorCount];
    }
    // 分配二维动态数组：存储TIP真空度参考值最大值
    double **referMaxVacuum = NULL; // 开辟动态数组行
    referMaxVacuum = new double *[paperCount]; // 开辟动态数组行
    for (int i = 0; i < paperCount; i++) // 开辟动态数组列
    {
        referMaxVacuum[i] = new double[factorCount];
    }

    // 创建工作簿句柄
    Book *bookRead = xlCreateXMLBook();
    // 注册 LibXL库
    bookRead->setKey("RenPeng", "windows-2228250808ceeb0a62b56669a4i6k6g3");

    // 设定各纸种抽气量计算TIP系数
    // 装载各纸种TIP抽气量系数工作簿
    if (bookRead->load("tipFactor.xlsx"))
    {
        // 获取工作簿中工作表数量
        int sheetNum = bookRead->sheetCount();
        // 装载工作表
        if (vFi == 798)
        {
            for (int i = 0; i < sheetNum; i++)
            {
                Sheet *sheet = bookRead->getSheet(i);
                // 定义：最大行数，最大列数，纸机车速[m/min]
                int lastRow, lastCol, reelSpeed;
                // 定义：定量范围[g/m2]，纸种
                string basisWeight, paperType;
                // 读取工作表数据
                lastCol = sheet->lastCol();
                lastRow = sheet->lastRow();
                if (sheet->readStr(2,1) != NULL)
                {
                    basisWeight = sheet->readStr(2, 1);
                }
                else
                {
                    basisWeight = "No limit";
                }
                if (sheet->readStr(0,1))
                {
                    paperType = sheet->readStr(0, 1);
                }
                else
                {
                    paperType = "No define";
                    cout << "[" << i << "]: " << "The type of Paper is not definded!" << endl;
                }
                reelSpeed = sheet->readNum(4, 1);
                // 定义：最低定量[g/m2]，最高定量[g/m2]
                double minWeight, maxWeight;
                // 从定量范围分离最低和最高定量
                if (basisWeight.find("-") != basisWeight.npos)
                {
                    minWeight = stod(basisWeight.substr(0, basisWeight.find("-")));
                    maxWeight = stod(basisWeight.substr(basisWeight.find("-") + 1, basisWeight.length()));
                }
                else
                {
                    minWeight = 0.0;
                    maxWeight = 0.0;
                    cout << "[" << i << "]: " << "Basisweight of the paper is: " << basisWeight << endl;
                }
                // 循环读取TIP真空抽气量系数
                for (int j = 0; j < factorCount; ++j)
                {
                    vacuumMinFactor[i][j] = sheet->readNum(j + 1, 8);
                    referMinVacuum[i][j]  = sheet->readNum(j + 1, 14);
                    vacuumMaxFactor[i][j] = sheet->readNum(j + 1, 9);
                    referMaxVacuum[i][j]  = sheet->readNum(j + 1, 15);
                }
            }
        }
        else
        {
            Sheet *sheet = bookRead->getSheet(vFi);
            // 定义：最大行数，最大列数，纸机车速[m/min]
            int lastRow, lastCol, reelSpeed;
            // 定义：定量范围[g/m2]，纸种
            string basisWeight, paperType;
            // 读取工作表数据
            lastCol     = sheet->lastCol();
            lastRow     = sheet->lastRow();
            basisWeight = sheet->readStr(2, 1);
            paperType   = sheet->readStr(0, 1);
            reelSpeed   = sheet->readNum(4, 1);
            // 判断是否与输入纸种相符
            if (paperType != paperTypeset)
            {
                cout << "\n"
                     << "********"
                     << "NOT MATCH!!!"
                     << "\t"
                     << "CHECK!!!" << endl;
            }
            // 定义：最低定量[g/m2]，最高定量[g/m2]
            double minWeight, maxWeight;
            // 从定量范围分离最低和最高定量
            minWeight = stod(basisWeight.substr(0, basisWeight.find("-")));
            maxWeight = stod(basisWeight.substr(basisWeight.find("-") + 1, basisWeight.length()));
            // 循环读取TIP真空抽气量系数
            for (int j = 0; j < factorCount; ++j)
            {
                vacuumMinFactor[vFi][j] = sheet->readNum(j + 1, 8);
                referMinVacuum[vFi][j]  = sheet->readNum(j + 1, 14);
                vacuumMaxFactor[vFi][j] = sheet->readNum(j + 1, 9);
                referMaxVacuum[vFi][j]  = sheet->readNum(j + 1, 15);
            }
        }
    }
    else
    {
        cout << bookRead->errorMessage() << endl;
    }
    // 储存TIP系数工作簿
    if (!bookRead->save("tipFactor.xlsx"))
    {
        cout << bookRead->errorMessage() << endl;
    }

    // 装载原始数据工作簿
    if (bookRead->load("originData.xlsx"))
    {
        // 定义：最大行数，最大列数，纸机车速[m/min]
        int lastRow, lastCol, reelSpeed;
        // 定义：最低定量[g/m2]，最高定量[g/m2]
        double minWeight, maxWeight;
        // 定义：定量范围[g/m2]，纸种
        string basisWeight, paperType;
        // 获取工作簿中工作表数量
        int sheetNum = bookRead->sheetCount();
        // 装载工作表
        for (int sheetNo = 0; sheetNo < sheetNum; sheetNo++)
        {
            // 循环装载工作表
            Sheet *sheet = bookRead->getSheet(sheetNo);
            // 读取工作表数据
            lastCol     = sheet->lastCol();
            lastRow     = sheet->lastRow();
            basisWeight = sheet->readStr(2, 1);
            paperType   = sheet->readStr(3, 1);
            reelSpeed   = sheet->readNum(1, 1);
            // 确定纸种系数的索引数
            try
            {
                vFi = paperID.at(paperType);
            }
            catch(const std::out_of_range& e)
            {
                std::cerr << e.what() << '\n'<< paperType << " was not found." <<std::endl;
            }
            // 从定量范围分离最低和最高定量
            minWeight = stod(basisWeight.substr(0, basisWeight.find("-")));
            maxWeight = stod(basisWeight.substr(basisWeight.find("-") + 1, basisWeight.length()));
            // 定义：抽吸位置数组，原始数据数组，抽汽流量[m3/min]
            const char *suckLocation[lastRow - 1];
            const char *suctionLocation[lastRow - 1];
            double originData[lastRow - 1][lastCol - 4];
            double unitMinAirflow[lastRow - 1];
            double unitMaxAirflow[lastRow - 1];

            // 控制数据表格式
            // 设置单元格格式，标题栏格式：字体尺寸14，垂直居中，水平居中，四周细边框
            Font *titleFont = bookRead->addFont();
            titleFont->setSize(14);
            Format *titleFormat = bookRead->addFormat();
            titleFormat->setAlignH(ALIGNH_CENTER);
            titleFormat->setAlignV(ALIGNV_CENTER);
            titleFormat->setFont(titleFont);
            titleFormat->setBorder(BORDERSTYLE_THIN);
            titleFormat->setWrap(true);
            // 设置单元格格式，标题栏格式：字体尺寸12，垂直居中，水平居中，四周细边框
            Font *textFont = bookRead->addFont();
            textFont->setSize(12);
            Format *textFormat = bookRead->addFormat();
            textFormat->setAlignH(ALIGNH_RIGHT);
            textFormat->setAlignV(ALIGNV_CENTER);
            textFormat->setFont(textFont);
            textFormat->setBorder(BORDERSTYLE_THIN);
            textFormat->setNumFormat(NUMFORMAT_NUMBER_D2);

            // 循环读取原始数据列表中的真空元件参数数据
            cout << "\n"
                 << "工作表[" << sheetNo + 1 << "]" << "：各真空抽吸元件所需抽气量：" << endl;
            for (int row = 1; row <= lastRow - 1; row++)
            {
                for (int col = startColnum; col <= lastCol - 1; col++)
                {
                    originData[row - 1][col - startColnum] = sheet->readNum(row, col);
                }
                // 读取抽吸位置
                suckLocation[row - 1] = sheet->readStr(row, startColnum - 1); // 读取列号和行号从0开始计算
                // 抽吸位置数组格式化
                string sLocation = formatString(suckLocation[row - 1]);
                // 将格式化字符串数组赋值
                suctionLocation[row - 1] = sLocation.c_str();
                // 进行抽气量计算，根据不同工段，选取不同公式
                if (originData[row - 1][0] < boxPart)
                {
                    // 长网纸机网部吸湿箱TIP计算公式
                    unitMinAirflow[row - 1] = originData[row - 1][1] / 1000.0 * originData[row - 1][4] \
                                         * vacuumMinFactor[vFi][int(originData[row - 1][0]) - 1];
                    unitMaxAirflow[row - 1] = originData[row - 1][1] / 1000.0 * originData[row - 1][4] \
                                         * vacuumMaxFactor[vFi][int(originData[row - 1][0]) - 1];
                    // 将公式写入Excel
                    if (processControl == 3)
                    {
                        // 写入标题栏
                        sheet->writeStr(0, writeColnum - 1, "referMinVacuum\n[SI]", titleFormat);
                        sheet->writeStr(0, writeColnum, "referMinVacuum\n[SI]", titleFormat);
                        sheet->writeStr(0, writeColnum + 1, "vacuumMinFactor\n[SI]", titleFormat);
                        sheet->writeStr(0, writeColnum + 2, "vacuumMaxFactor\n[SI]", titleFormat);
                        sheet->writeStr(0, writeColnum + 3, "airMinFlow\n[m3/min]", titleFormat);
                        sheet->writeStr(0, writeColnum + 4, "airMaxFlow\n[m3/min]", titleFormat);
                        // 写入TIP参考真空度
                        sheet->writeNum(row, writeColnum - 1, \
                                        referMinVacuum[vFi][int(originData[row - 1][0]) - 1], textFormat);
                        sheet->writeNum(row, writeColnum, \
                                        referMaxVacuum[vFi][int(originData[row - 1][0]) - 1], textFormat);
                        // 写入TIP系数
                        sheet->writeNum(row, writeColnum + 1, \
                                        vacuumMinFactor[vFi][int(originData[row - 1][0]) - 1], textFormat);
                        sheet->writeNum(row, writeColnum + 2, \
                                        vacuumMaxFactor[vFi][int(originData[row - 1][0]) - 1], textFormat);
                        // 定义计算使用的单元格行数和列字母
                        const char colVocabulary_1  = 65 + 4 + 1; // F列
                        const char colVocabulary_4  = 65 + 4 + 4; // I列
                        const char colVocabulary_10 = 65 + 4 + 10; // O列
                        const char colVocabulary_11 = 65 + 4 + 11; // P列
                        string rowStr = to_string(row + 1);
                        string cV1, cV4, cV10, cV11;
                        char tranStr[2] = {};
                        // 获取计算使用的第一个单元格编号
                        tranStr[0] = colVocabulary_1;
                        cV1 = tranStr;
                        string cellNo_1 = cV1 + rowStr;
                        // 获取计算使用的第二个单元格编号
                        tranStr[0] = colVocabulary_4;
                        cV4 = tranStr;
                        string cellNo_4 = cV4 + rowStr;
                        // 获取计算使用的第三个单元格编号，最小值
                        tranStr[0] = colVocabulary_10;
                        cV10 = tranStr;
                        string cellNo_10 = cV10 + rowStr;
                        // 获取计算使用的第三个单元格编号，最小值
                        tranStr[0] = colVocabulary_11;
                        cV11 = tranStr;
                        string cellNo_11 = cV11 + rowStr;
                        // 获得计算公式字符串
                        string cellCalmin = cellNo_1 + "/1000*" + cellNo_4 + "*" + cellNo_10;
                        string cellCalmax = cellNo_1 + "/1000*" + cellNo_4 + "*" + cellNo_11;
                        // 写入公式
                        sheet->writeFormula(row, writeColnum + 3, cellCalmin.c_str(), textFormat);
                        sheet->writeFormula(row, writeColnum + 4, cellCalmax.c_str(), textFormat);
                        // cout << cellCal << endl;
                    }
                }
                else
                {
                    if (originData[row - 1][0] > boxPart && originData[row - 1][0] < rollPartstart \
                                                         || originData[row - 1][0] > rollPartstop)
                    {
                        // 长网纸机网部低真空吸水箱、高真空吸水箱、毛毯吸水箱的TIP计算公式
                        unitMinAirflow[row - 1] = originData[row - 1][1] / 1000.0 * originData[row - 1][2] \
                                                * originData[row - 1][3] / 1000.0 * originData[row - 1][4] \
                                                * vacuumMinFactor[vFi][int(originData[row - 1][0]) - 1];
                        unitMaxAirflow[row - 1] = originData[row - 1][1] / 1000.0 * originData[row - 1][2] \
                                                * originData[row - 1][3] / 1000.0 * originData[row - 1][4] \
                                                * vacuumMaxFactor[vFi][int(originData[row - 1][0]) - 1];
                        // 将公式写入Excel
                        if (processControl == 3)
                        {
                            // 写入标题栏
                            sheet->writeStr(0, writeColnum - 1, "referMinVacuum\n[SI]", titleFormat);
                            sheet->writeStr(0, writeColnum, "referMinVacuum\n[SI]", titleFormat);
                            sheet->writeStr(0, writeColnum + 1, "vacuumMinFactor\n[SI]", titleFormat);
                            sheet->writeStr(0, writeColnum + 2, "vacuumMaxFactor\n[SI]", titleFormat);
                            sheet->writeStr(0, writeColnum + 3, "airMinFlow\n[m3/min]", titleFormat);
                            sheet->writeStr(0, writeColnum + 4, "airMaxFlow\n[m3/min]", titleFormat);
                            // 写入TIP参考真空度
                            sheet->writeNum(row, writeColnum - 1, \
                                            referMinVacuum[vFi][int(originData[row - 1][0]) - 1], textFormat);
                            sheet->writeNum(row, writeColnum, \
                                            referMaxVacuum[vFi][int(originData[row - 1][0]) - 1], textFormat);
                            // 写入TIP系数
                            sheet->writeNum(row, writeColnum + 1, \
                                            vacuumMinFactor[vFi][int(originData[row - 1][0]) - 1], textFormat);
                            sheet->writeNum(row, writeColnum + 2, \
                                            vacuumMaxFactor[vFi][int(originData[row - 1][0]) - 1], textFormat);
                            // 定义计算使用的单元格行数和列字母
                            const char colVocabulary_1  = 65 + 4 + 1; // F列
                            const char colVocabulary_2  = 65 + 4 + 2; // G列
                            const char colVocabulary_3  = 65 + 4 + 3; // H列
                            const char colVocabulary_4  = 65 + 4 + 4; // I列
                            const char colVocabulary_10 = 65 + 4 + 10; // O列
                            const char colVocabulary_11 = 65 + 4 + 11; // P列
                            string rowStr = to_string(row + 1);
                            string cV1, cV2, cV3, cV4, cV10, cV11;
                            char tranStr[2] = {};
                            // 获取计算使用的第一个单元格编号
                            tranStr[0] = colVocabulary_1;
                            cV1 = tranStr;
                            string cellNo_1 = cV1 + rowStr;
                            // 获取计算使用的第二个单元格编号
                            tranStr[0] = colVocabulary_2;
                            cV2 = tranStr;
                            string cellNo_2 = cV2 + rowStr;
                            // 获取计算使用的第三个单元格编号
                            tranStr[0] = colVocabulary_3;
                            cV3 = tranStr;
                            string cellNo_3 = cV3 + rowStr;
                            // 获取计算使用的第四个单元格编号
                            tranStr[0] = colVocabulary_4;
                            cV4 = tranStr;
                            string cellNo_4 = cV4 + rowStr;
                            // 获取计算使用的第五个单元格编号，最小值
                            tranStr[0] = colVocabulary_10;
                            cV10 = tranStr;
                            string cellNo_10 = cV10 + rowStr;
                            // 获取计算使用的第五个单元格编号，最大值
                            tranStr[0] = colVocabulary_11;
                            cV11 = tranStr;
                            string cellNo_11 = cV11 + rowStr;
                            // 获得计算公式字符串
                            string cellCalmin = cellNo_1 + "/1000*" + cellNo_2 + "*" + cellNo_3 + "/1000*" \
                                              + cellNo_4 + "*" + cellNo_10;
                            string cellCalmax = cellNo_1 + "/1000*" + cellNo_2 + "*" + cellNo_3 + "/1000*" \
                                              + cellNo_4 + "*" + cellNo_11;
                            // 写入公式
                            sheet->writeFormula(row, writeColnum + 3, cellCalmin.c_str(), textFormat);
                            sheet->writeFormula(row, writeColnum + 4, cellCalmax.c_str(), textFormat);
                            // cout << cellCal << endl;
                        }
                    }
                    else
                    {
                        if (originData[row - 1][3] <= 25.4 * 6)
                        {
                            // 长网纸机网部伏辊、吸移辊、压榨辊的TIP计算公式（辊径小于6英寸，即152.4mm)
                            unitMinAirflow[row - 1] = originData[row - 1][1] / 1000.0 * \
                                                      originData[row - 1][3] / 1000.0 * \
                                                      vacuumMinFactor[vFi][int(originData[row - 1][0]) - 1] * \
                                                      originData[row - 1][4];
                            unitMaxAirflow[row - 1] = originData[row - 1][1] / 1000.0 * \
                                                      originData[row - 1][3] / 1000.0 * \
                                                      vacuumMaxFactor[vFi][int(originData[row - 1][0]) - 1] * \
                                                      originData[row - 1][4];
                            // 将公式写入Excel
                            if (processControl == 3)
                            {
                                // 写入标题栏
                                sheet->writeStr(0, writeColnum - 1, "referMinVacuum\n[SI]", titleFormat);
                                sheet->writeStr(0, writeColnum, "referMinVacuum\n[SI]", titleFormat);
                                sheet->writeStr(0, writeColnum + 1, "vacuumMinFactor\n[SI]", titleFormat);
                                sheet->writeStr(0, writeColnum + 2, "vacuumMaxFactor\n[SI]", titleFormat);
                                sheet->writeStr(0, writeColnum + 3, "airMinFlow\n[m3/min]", titleFormat);
                                sheet->writeStr(0, writeColnum + 4, "airMaxFlow\n[m3/min]", titleFormat);
                                // 写入TIP参考真空度
                                sheet->writeNum(row, writeColnum - 1, \
                                                referMinVacuum[vFi][int(originData[row - 1][0]) - 1], textFormat);
                                sheet->writeNum(row, writeColnum, \
                                                referMaxVacuum[vFi][int(originData[row - 1][0]) - 1], textFormat);
                                // 写入TIP系数
                                sheet->writeNum(row, writeColnum + 1, \
                                                vacuumMinFactor[vFi][int(originData[row - 1][0]) - 1], textFormat);
                                sheet->writeNum(row, writeColnum + 2, \
                                                vacuumMaxFactor[vFi][int(originData[row - 1][0]) - 1], textFormat);
                                // 定义计算使用的单元格行数和列字母
                                const char colVocabulary_1  = 65 + 4 + 1; // F列
                                const char colVocabulary_2  = 65 + 4 + 2; // G列
                                const char colVocabulary_3  = 65 + 4 + 3; // H列
                                const char colVocabulary_4  = 65 + 4 + 4; // I列
                                const char colVocabulary_10 = 65 + 4 + 10; // O列
                                const char colVocabulary_11 = 65 + 4 + 11; // P列
                                string rowStr = to_string(row + 1);
                                string cV1, cV3, cV4, cV10, cV11;
                                char tranStr[2] = {};
                                // 获取计算使用的第一个单元格编号
                                tranStr[0] = colVocabulary_1;
                                cV1 = tranStr;
                                string cellNo_1 = cV1 + rowStr;
                                // 获取计算使用的第二个单元格编号
                                tranStr[0] = colVocabulary_3;
                                cV3 = tranStr;
                                string cellNo_3 = cV3 + rowStr;
                                // 获取计算使用的第三个单元格编号
                                tranStr[0] = colVocabulary_4;
                                cV4 = tranStr;
                                string cellNo_4 = cV4 + rowStr;
                                // 获取计算使用的第四个单元格编号，最小值
                                tranStr[0] = colVocabulary_10;
                                cV10 = tranStr;
                                string cellNo_10 = cV10 + rowStr;
                                // 获取计算使用的第四个单元格编号，最大值
                                tranStr[0] = colVocabulary_11;
                                cV11 = tranStr;
                                string cellNo_11 = cV11 + rowStr;
                                // 获得计算公式字符串
                                string cellCalmin = cellNo_1 + "/1000*" + cellNo_3 + "/1000*" + cellNo_10 + "*" + cellNo_4;
                                string cellCalmax = cellNo_1 + "/1000*" + cellNo_3 + "/1000*" + cellNo_11 + "*" + cellNo_4;
                                // 写入公式
                                sheet->writeFormula(row, writeColnum + 3, cellCalmin.c_str(), textFormat);
                                sheet->writeFormula(row, writeColnum + 4, cellCalmax.c_str(), textFormat);
                                // cout << cellCal << endl;
                            }
                        }
                        else
                        {
                            // 根据辊径、弦长计算弧长
                            double r = originData[row - 1][7] / 2.0;
                            double chordLength = originData[row - 1][3];
                            double arcLength = r * asin(1.0 - 0.5 * pow(chordLength / r, 2));
                            // 长网纸机网部伏辊、吸移辊、压榨辊的TIP计算公式（辊径大于6英寸，即152.4mm)
                            unitMinAirflow[row - 1] = originData[row - 1][1] / 1000.0 * arcLength / 1000.0 \
                                                    * vacuumMinFactor[vFi][int(originData[row - 1][0]) - 1];
                            unitMaxAirflow[row - 1] = originData[row - 1][1] / 1000.0 * arcLength / 1000.0 \
                                                    * vacuumMaxFactor[vFi][int(originData[row - 1][0]) - 1];
                            // 将公式写入Excel
                            if (processControl == 3)
                            {
                                // 写入标题栏
                                sheet->writeStr(0, writeColnum - 1, "referMinVacuum\n[SI]", titleFormat);
                                sheet->writeStr(0, writeColnum, "referMinVacuum\n[SI]", titleFormat);
                                sheet->writeStr(0, writeColnum + 1, "vacuumMinFactor\n[SI]", titleFormat);
                                sheet->writeStr(0, writeColnum + 2, "vacuumMaxFactor\n[SI]", titleFormat);
                                sheet->writeStr(0, writeColnum + 3, "airMinFlow\n[m3/min]", titleFormat);
                                sheet->writeStr(0, writeColnum + 4, "airMaxFlow\n[m3/min]", titleFormat);
                                // 写入TIP参考真空度
                                sheet->writeNum(row, writeColnum - 1, \
                                                referMinVacuum[vFi][int(originData[row - 1][0]) - 1], textFormat);
                                sheet->writeNum(row, writeColnum, \
                                                referMaxVacuum[vFi][int(originData[row - 1][0]) - 1], textFormat);
                                // 写入TIP系数
                                sheet->writeNum(row, writeColnum - 1, \
                                                vacuumMinFactor[vFi][int(originData[row - 1][0]) - 1], textFormat);
                                sheet->writeNum(row, writeColnum, \
                                                vacuumMaxFactor[vFi][int(originData[row - 1][0]) - 1], textFormat);
                                // 定义计算使用的单元格行数和列字母
                                const char colVocabulary_1  = 65 + 4 + 1; // F
                                const char colVocabulary_3  = 65 + 4 + 3; // H
                                const char colVocabulary_4  = 65 + 4 + 4; // I
                                const char colVocabulary_7  = 65 + 4 + 7; // L
                                const char colVocabulary_10 = 65 + 4 + 10; // O
                                const char colVocabulary_11 = 65 + 4 + 11; // P
                                string rowStr = to_string(row + 1);
                                string cV1, cV3, cV4, cV7, cV10, cV11;
                                char tranStr[2] = {};
                                // 获取计算使用的第一个单元格编号
                                tranStr[0] = colVocabulary_1;
                                cV1 = tranStr;
                                string cellNo_1 = cV1 + rowStr;
                                // 获取计算使用的第二个单元格编号
                                tranStr[0] = colVocabulary_3;
                                cV3 = tranStr;
                                string cellNo_3 = cV3 + rowStr;
                                // 获取计算使用的第三个单元格编号
                                tranStr[0] = colVocabulary_4;
                                cV4 = tranStr;
                                string cellNo_4 = cV4 + rowStr;
                                // 获取计算使用的第四个单元格编号
                                tranStr[0] = colVocabulary_7;
                                cV7 = tranStr;
                                string cellNo_7 = cV7 + rowStr;
                                // 获取计算使用的第五个单元格编号，最小系数
                                tranStr[0] = colVocabulary_10;
                                cV10 = tranStr;
                                // 获取计算使用的第五个单元格编号，最大系数
                                tranStr[0] = colVocabulary_10;
                                cV11 = tranStr;
                                string cellNo_10 = cV10 + rowStr;
                                string cellNo_11 = cV11 + rowStr;
                                // 获得计算公式字符串
                                string cellCalmin = cellNo_1 + "/1000*" + cellNo_7 + "/2*asin(1-0.5*power(" + cellNo_3 \
                                                  + "/" + cellNo_7 + "*2,2))" + "/1000*" + cellNo_10 + "*" + cellNo_4;
                                string cellCalmax = cellNo_1 + "/1000*" + cellNo_7 + "/2*asin(1-0.5*power(" + cellNo_3 \
                                                  + "/" + cellNo_7 + "*2,2))" + "/1000*" + cellNo_11 + "*" + cellNo_4;
                                // 写入公式，c_str()将字符串（string）转换为（char）
                                sheet->writeFormula(row, writeColnum + 3, cellCalmin.c_str(), textFormat);
                                sheet->writeFormula(row, writeColnum + 4, cellCalmax.c_str(), textFormat);
                                // cout << cellCal << endl;
                            }
                        }
                    }
                }
                // 输出不同抽吸位置计算的抽气量至屏幕
                cout << suctionLocation[row - 1] << "\t" << unitMaxAirflow[row - 1] << endl;
                // 将计算数据结果写入Excel
                if (processControl == 2)
                {
                    sheet->writeStr(0, writeColnum - 1, "referMinVacuum\n[SI]", titleFormat);
                    sheet->writeStr(0, writeColnum, "referMinVacuum\n[SI]", titleFormat);
                    sheet->writeStr(0, writeColnum + 1, "vacuumMinFactor\n[SI]", titleFormat);
                    sheet->writeStr(0, writeColnum + 2, "vacuumMaxFactor\n[SI]", titleFormat);
                    sheet->writeStr(0, writeColnum + 3, "airMinFlow\n[m3/min]", titleFormat);
                    sheet->writeStr(0, writeColnum + 4, "airMaxFlow\n[m3/min]", titleFormat);
                    sheet->writeNum(row, writeColnum - 1, \
                                    referMinVacuum[vFi][int(originData[row - 1][0]) - 1], textFormat);
                    sheet->writeNum(row, writeColnum, \
                                    referMaxVacuum[vFi][int(originData[row - 1][0]) - 1], textFormat);
                    sheet->writeNum(row, writeColnum + 1, \
                                    vacuumMinFactor[vFi][int(originData[row - 1][0]) - 1], textFormat);
                    sheet->writeNum(row, writeColnum + 2, \
                                    vacuumMaxFactor[vFi][int(originData[row - 1][0]) - 1], textFormat);
                    sheet->writeNum(row, writeColnum + 3, unitMinAirflow[row - 1], textFormat);
                    sheet->writeNum(row, writeColnum + 4, unitMaxAirflow[row - 1], textFormat);
                }
            }
        }

    }
    else
    {
        // 装载数据表格错误，输出错误信息
        cout << bookRead->errorMessage() << endl;
    }

    // 储存数据表格
    if (!bookRead->save("originData.xlsx"))
    {
        // 存储失败输出错误信息
        cout << bookRead->errorMessage() << endl;
    }
    // 释放通道
    bookRead->release();

    //释放二维数组
    for(int i = 0; i < 1; ++i)
    {
        vacuumMinFactor[i] = NULL;
        delete[] vacuumMinFactor[i];
    }
    delete[] vacuumMinFactor;
    vacuumMinFactor = NULL;
    for(int i = 0; i < paperCount; ++i)
    {
        referMinVacuum[i] = NULL;
        delete[] referMinVacuum[i];
    }
    delete[] referMinVacuum;
    referMinVacuum = NULL;

    //释放二维数组
    for(int i = 0; i < 1; ++i)
    {
        vacuumMaxFactor[i] = NULL;
        delete[] vacuumMaxFactor[i];
    }
    delete[] vacuumMaxFactor;
    vacuumMaxFactor = NULL;

    for(int i = 0; i < paperCount; ++i)
    {
        referMaxVacuum[i] = NULL;
        delete[] referMaxVacuum[i];
    }
    delete[] referMaxVacuum;
    referMaxVacuum = NULL;

    // 输出结束提示
    cout << "\nPrint any key to continue..." << endl;
    _getch();
    return 0;
}
