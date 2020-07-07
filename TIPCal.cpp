// Copyright 2008...
// License(BSD/GPL...)
// Author: RenPeng
// This is
#include <iostream>
//#include <cstdlib>
#include <algorithm>
#include <cmath>
#include <map>
#include <conio.h>
#include "libxl.h"
#include "rpfunc.hpp"

// 设定TIP系数纸种类型和系数个数
#define paperCount 25
#define factorCount 30

using namespace std;
using namespace libxl;
using namespace basicmath;
using namespace basicformat;
using namespace calconfig;


// 计算主程序
int main()
{
  // 控制台显示乱码纠正
  // system("chcp 65001"); //设置字符集，最终编译时加上，否则调试时出问题
  // 纸种系数字典paperID定义在rpfunc.hpp中

  // 定义数据表中：真空元件信息起始列号、写入数据起始列号、纸机真空元件划分的段数
  int startColnum, writeColnum, elementNum;
  startColnum =  4; // 真空元件抽吸位置名称列开始
  writeColnum = 13; // M列开始
  elementNum  = 25; // unitLocation数值范围1~25，意思是纸机真空元件划分了25种不同类型
  // 定义计算公式区分位置的工段号
  Papermachine PM0;
  PM0.setVacuumConfig(4, 5, 6, 21, 22);
  // PM0.foilStop......吸湿箱最大工段号，1~3全部使用粗算公式，并不考虑实际暴露面积，至考虑元件数量，真空度低，无需校准
  // PM0.flatStop......真空箱最大工段号，4~5，分别为低真空和高真空吸水箱，4之后是水线，此时计算公式考虑暴露面积
  // PM0.rollBegin.....真空吸辊开始工段号，6~21，分别为伏辊低真空、高真空，吸移辊低、高真空，压榨辊低真空、高真空
  // PM0.rollStop......真空吸辊结束工段号
  // PM0.feltBegin.....毛毯吸水箱和转移箱开始工段号，22~25
  // 定义纸种类型对应的数组索引标号
  int vFi = 798;
  // 定义：需要计算的纸种类型
  string paperTypeset;

  // 定义纸种，确定计算数据，输出屏幕说明
  printHeadClaim();

  // 判断读取数据是否正确
  while (1)
  {
    // 定义临时字符数组用来存取屏幕读入的数据
    char pTtemp[6] = "798";
    std::cin.getline(pTtemp, 6);
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
      std::cout << "输入有误，请重新输入：" << endl;
    }
    else
    {
      if (pT == "")
      {
        paperTypeset = "MULTIPLE";
        break;
      }
      else
      {
        vFi = stoi(pT) - 1;
        for (std::map<string, int>::iterator item = paperID.begin(); item != paperID.end(); item++)
        {
          if (item->second == vFi)
            {
              paperTypeset = item->first;
              std::cout << "*************************************************" << endl;
              std::cout << "Calculating paper type: " << paperTypeset << endl;
              std::cout << "*************************************************" << endl;
            }
        }
        break;
      }
    }
  }
  int processControl = 3;
  std::cout << "[1] 临时计算，计算值只显示在屏幕上；\n"
       << "[2] 存储结果，计算值存储到原Excel；\n"
       << "[3] 存储公式，计算公式写入Excel。" << endl;
  std::cout << "请输入进程参数（默认值3）：";
  while (1)
  {
    // 定义临时字符数组用来存取屏幕读入的数据
    char pCtemp[6] = "3";
    std::cin.getline(pCtemp, 6);
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
      std::cout << "输入有误，请重新输入：" << endl;
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

  // 定义标准定量
  double stdMinWeight, stdMaxWeight;

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
    // 如果纸种选择输入项为默认值，则循环读取所有纸种的系数列表
    if (vFi == 798)
    {
      for (int i = 0; i < sheetNum; i++)
      {
        // 装载工作表
        Sheet *sheet = bookRead->getSheet(i);
        // 定义：最大行数，最大列数，纸机车速[m/min]
        int lastRow, lastCol, reelSpeed;
        // 定义：定量范围[g/m2]，纸种
        string basisWeight, paperType;
        // 读取工作表数据的行数和列数
        lastCol = sheet->lastCol();
        lastRow = sheet->lastRow();
        // 读取工作表中纸张定量
        if (sheet->readStr(3,1) != NULL)
        {
          basisWeight = sheet->readStr(3, 1);
        }
        // 若定量单元格无数据，输出定量无限制的信息
        else
        {
          basisWeight = "No limit";
        }
        // 读取工作表中纸张类型
        if (sheet->readStr(0,1))
        {
          paperType = sheet->readStr(0, 1);
        }
        // 若纸张类型单元格无数据，输出纸种未定义信息
        else
        {
          paperType = "No define";
          std::cout << "!--" << i << "--!: "
               << "The type of Paper is not definded!" << endl;
        }
        // 读取工作表中纸机车速
        reelSpeed = sheet->readNum(4, 1);
        // 定义：最低定量[g/m2]，最高定量[g/m2]
        double minWeight, maxWeight;
        // 从定量范围分离最低和最高定量
        if (basisWeight.find("-") != basisWeight.npos)
        {
          minWeight = stod(basisWeight.substr(0, basisWeight.find("-")));
          maxWeight = stod(basisWeight.substr(basisWeight.find("-") + 1, basisWeight.length()));
        }
        // 若定量中没有"-"，则判断没有定量限制
        else
        {
          minWeight = 0.0;
          maxWeight = 0.0;
          std::cout << "!--" << i << "--!: "
               << "Basisweight of the paper is: " << basisWeight << endl;
        }
        // 将系数列表中的标准定量赋值给全局变量
        stdMinWeight = minWeight;
        stdMaxWeight = maxWeight;
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
      // 读取输入的纸种对应的系数表格
      Sheet *sheet = bookRead->getSheet(vFi);
      // 定义：最大行数，最大列数，纸机车速[m/min]
      int lastRow, lastCol, reelSpeed;
      // 定义：定量范围[g/m2]，纸种
      string basisWeight, paperType;
      // 读取工作表数据
      lastCol     = sheet->lastCol();
      lastRow     = sheet->lastRow();
      basisWeight = sheet->readStr(3, 1);
      paperType   = sheet->readStr(1, 1);
      reelSpeed   = sheet->readNum(4, 1);
      // 判断是否与输入纸种相符
      if (paperType != paperTypeset)
      {
        std::cout << "\n"
             << "*************************************************"
             << "\n"
             << "NOT MATCH!!!"
             << " "
             << "CHECK FACTOR DATA!!!"
             << "\n"
             << "TIP FACTOR USING IS: "
             << "\t" << paperType << "\n"
             << "*************************************************" << endl;
        std::cin.clear(); // 清除流的错误标记
        std::cin.ignore( numeric_limits<streamsize>::max(), '\n' ); // 清空输入流
        std::cin.get(); // 等待用户输入回车
      }
      // 定义：最低定量[g/m2]，最高定量[g/m2]
      double minWeight, maxWeight;
      // 从定量范围分离最低和最高定量
      if (basisWeight.find("-") != basisWeight.npos)
      {
        minWeight = stod(basisWeight.substr(0, basisWeight.find("-")));
        maxWeight = stod(basisWeight.substr(basisWeight.find("-") + 1, basisWeight.length()));
      }
      // 若定量中没有"-"，则判断没有定量限制
      else
      {
        minWeight = 0.0;
        maxWeight = 0.0;
        std::cout << "!--" << paperType << "--!: "
             << "Basisweight of the paper is: " << basisWeight << endl;
      }
      // 将系数列表中的标准定量赋值给全局变量
      stdMinWeight = minWeight;
      stdMaxWeight = maxWeight;
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
    std::cout << bookRead->errorMessage() << endl;
  }
  // 储存TIP系数工作簿
  if (!bookRead->save("tipFactor.xlsx"))
  {
    std::cout << bookRead->errorMessage() << endl;
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
    // 定义：最后一行第一个字符（用于判断是否有统计行）
    string lastRowVal = "NULL";
    // 获取工作簿中工作表数量
    int sheetNum = bookRead->sheetCount();
    // 定义：是否存在统计行
    bool conCal = false;

    // 装载工作表
    for (int sheetNo = 0; sheetNo < sheetNum; sheetNo++)
    {
      // 循环装载工作表
      Sheet *sheet = bookRead->getSheet(sheetNo);
      // 读取工作表数据
      // 读取工作表行列数和基本信息
      lastCol     = sheet->lastCol();
      lastRow     = sheet->lastRow();
      basisWeight = sheet->readStr(2, 1);
      paperType   = sheet->readStr(3, 1);
      reelSpeed   = sheet->readNum(1, 1);
      // 判断数据表中需要计算纸种是否与之前选择计算纸种相符
      if (paperType != paperTypeset && paperTypeset != "MULTIPLE")
      {
        std::cout << "\n"
             << "*************************************************" << "\n"
             << "NOT MATCH!!!"
             << "\t"
             << "CHECK ORIGINAL DATA!!!"
             << "\n"
             << "PAPER TYPE IN ORIGINAL DATA IS NOT MATCH WITH PAPER TYPE SET!"
             << "\n"
             << "*************************************************" << endl;
        std::cin.clear(); // 清除流的错误标记
        std::cin.ignore( numeric_limits<streamsize>::max(), '\n' ); // 清空输入流
        std::cin.get(); // 等待用户输入回车
      }
      // 判断是否已经有计算结果列，若有读取最后一行的值
      if (sheet->cellType(0, writeColnum + 2) != CELLTYPE_EMPTY)
      {
        lastRowVal  = sheet->readStr(lastRow - 1, writeColnum + 2);
      }
      // 判断是否有统计行
      if (lastRowVal == "Total")
      {
        lastRow = lastRow - 1;
        conCal = true;
      }

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
      if (basisWeight.find("-") != basisWeight.npos)
      {
        minWeight = stod(basisWeight.substr(0, basisWeight.find("-")));
        maxWeight = stod(basisWeight.substr(basisWeight.find("-") + 1, basisWeight.length()));
      }
      // 若定量中没有"-"，则判断没有定量限制
      else
      {
        minWeight = 0.0;
        maxWeight = 0.0;
        std::cout << "!--" << paperType << "--!: "
             << "Basisweight of the paper is: " << basisWeight << endl;
      }
      // 定义：抽吸位置数组，原始数据数组，抽气流量[m3/min]
      const char *suckLocation[lastRow - 1];
      const char *suctionLocation[lastRow - 1];
      double originData[lastRow - 1][lastCol - 4];
      double unitMinAirflow[lastRow - 1];
      double unitMaxAirflow[lastRow - 1];
      // 定义最小和最大总抽气量变量
      double totalMinAirflow, totalMaxAirflow;
      totalMinAirflow = 0.0;
      totalMaxAirflow = 0.0;
      // 定义修正后最小和最大总抽气量变量
      double totalCrtMinFlow, totalCrtMaxFlow;
      totalCrtMinFlow = 0.0;
      totalCrtMaxFlow = 0.0;
      // 定义真空暴露时间存储数组，单缝时间，元件时间，工段时间
      double slotVacTime[lastRow - 1], unitVacTime[lastRow - 1], partVacTime[lastRow - 1];
      // 定义修正抽气量，真空度修正抽气量和定量修正抽气量
      double crtMinFlow[lastRow - 1], crtMaxFlow[lastRow - 1];

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
      titleFormat->setFillPattern(FILLPATTERN_GRAY6P25);
      titleFormat->setShrinkToFit(1);
      titleFormat->setPatternBackgroundColor(COLOR_TURQUOISE_CL);
      titleFormat->setPatternForegroundColor(COLOR_TURQUOISE_CL);
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
      std::cout << "\n"
           << "工作表[" << sheetNo + 1 << "]" << "：各真空抽吸元件所需抽气量：" << endl;
      std::cout << formatString("------Suction Location","-") << "\t"
           << formatString("minFlow","-",12)
           << formatString("maxFlow","-",12)
           << formatString("slotTime", "-",12)
           << formatString("unitTime", "-",12)
           << formatString("partTime", "-",12)
           << formatString("fixMinFlow", "-",12)
           << formatString("fixMaxFlow", "-",12)
           << endl;
      for (int row = 1; row <= lastRow - 1; row++)
      {
        for (int col = startColnum; col <= lastCol - 1; col++)
        {
          originData[row - 1][col - startColnum] = sheet->readNum(row, col);
        }
        // 读取抽吸位置，此时读取的是字符，将字符存入字符数组
        suckLocation[row - 1] = sheet->readStr(row, startColnum - 1); // 读取列号和行号从0开始计算
        // 抽吸位置数组格式化，将字符格式化成字符串
        string sLocation = formatString(suckLocation[row - 1], " ", 56);
        // 将格式化字符串数组赋值，将字符串再转化成字符
        suctionLocation[row - 1] = sLocation.c_str();
        // 进行抽气量计算，根据不同工段，选取不同公式
        if (originData[row - 1][0] < PM0.foilStop)
        {
          // 长网纸机网部吸湿箱抽气量TIP计算公式
          unitMinAirflow[row - 1] = originData[row - 1][1] / 1000.0 * originData[row - 1][4] \
                               * vacuumMinFactor[vFi][int(originData[row - 1][0]) - 1];
          unitMaxAirflow[row - 1] = originData[row - 1][1] / 1000.0 * originData[row - 1][4] \
                               * vacuumMaxFactor[vFi][int(originData[row - 1][0]) - 1];
          // 修正抽气量计算公式
          crtMinFlow[row - 1] = vacFixFlow(referMinVacuum[vFi][int(originData[row - 1][0]) - 1], originData[row - 1][5]) \
                                 * unitMinAirflow[row - 1] * weightFixFlow(stdMaxWeight, maxWeight); // 数据J列
          crtMaxFlow[row - 1] = vacFixFlow(referMaxVacuum[vFi][int(originData[row - 1][0]) - 1], originData[row - 1][6]) \
                                 * unitMaxAirflow[row - 1] * weightFixFlow(stdMinWeight, minWeight); // 数据K列
          // 长网纸机真空暴露时间计算公式
          slotVacTime[row - 1] = originData[row - 1][3] / reelSpeed * 60.0;
          unitVacTime[row - 1] = originData[row - 1][3] * originData[row - 1][2] / reelSpeed * 60.0;
          partVacTime[row - 1] = originData[row - 1][3] * originData[row - 1][2] / reelSpeed * 60.0 \
                               * originData[row - 1][4];
          // 将公式写入Excel
          if (processControl == 3)
          {
            // 写入标题栏
            sheet->writeStr(0, writeColnum - 1, "refMinVac\n[SI]", titleFormat);
            sheet->writeStr(0, writeColnum    , "refMaxVac\n[SI]", titleFormat);
            sheet->writeStr(0, writeColnum + 1, "minFactor\n[SI]", titleFormat);
            sheet->writeStr(0, writeColnum + 2, "maxFactor\n[SI]", titleFormat);
            sheet->writeStr(0, writeColnum + 3, "airMinFlow\n[m3/min]", titleFormat);
            sheet->writeStr(0, writeColnum + 4, "airMaxFlow\n[m3/min]", titleFormat);
            sheet->writeStr(0, writeColnum + 5, "slotVacTime\n[ms]", titleFormat);
            sheet->writeStr(0, writeColnum + 6, "unitVacTime\n[ms]", titleFormat);
            sheet->writeStr(0, writeColnum + 7, "partVacTime\n[ms]", titleFormat);
            sheet->writeStr(0, writeColnum + 8, "fixMinFlow\n[m3/min]", titleFormat);
            sheet->writeStr(0, writeColnum + 9, "fixMaxFlow\n[m3/min]", titleFormat);
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
            // 定义抽气量计算使用的单元格行数和列字母
            string rowStr = to_string(row + 1);
            string cellNo_1, cellNo_2, cellNo_3, cellNo_4, cellNo_10, cellNo_11;
            string cellNo_12, cellNo_13;
            // 获取抽气量计算使用的单元格编号
            cellNo_1  = transToVocabulary(5)  + rowStr; // F列
            cellNo_4  = transToVocabulary(8)  + rowStr; // I列
            cellNo_10 = transToVocabulary(14) + rowStr; // O列 （最小值）
            cellNo_11 = transToVocabulary(15) + rowStr; // P列 （最大值）
            // 获取真空暴露时间计算使用的单元格编号
            cellNo_3  = transToVocabulary(7)  + rowStr; // H列
            cellNo_2  = transToVocabulary(6)  + rowStr; // G列
            // 获取修正系数计算使用单元格编号
            cellNo_12 = transToVocabulary(16) + rowStr; // Q列
            cellNo_13 = transToVocabulary(17) + rowStr; // R列
            // 获得抽气量计算公式字符串
            string cellCalmin = cellNo_1 + "/1000*" + cellNo_4 + "*" + cellNo_10;
            string cellCalmax = cellNo_1 + "/1000*" + cellNo_4 + "*" + cellNo_11;
            // 获得修正抽气量计算公式字符串
            string cellFixmin = to_string(vacFixFlow(referMinVacuum[vFi][int(originData[row - 1][0]) - 1], originData[row - 1][5])) \
                              + "*" + to_string(weightFixFlow(stdMaxWeight, maxWeight)) + "*" +cellNo_12;
            string cellFixmax = to_string(vacFixFlow(referMaxVacuum[vFi][int(originData[row - 1][0]) - 1], originData[row - 1][6])) \
                              + "*" + to_string(weightFixFlow(stdMinWeight, minWeight)) + "*" +cellNo_13;
            // 获得真空暴露时间计算公式字符串
            string timeCalSlot = cellNo_3 + "/" + to_string(reelSpeed) + "*60";
            string timeCalUnit = cellNo_3 + "/" + to_string(reelSpeed) + "*60" + "*" + cellNo_2;
            string timeCalPart = cellNo_3 + "/" + to_string(reelSpeed) + "*60" + "*" + cellNo_2 \
                               + "*" + cellNo_4;
            // 写入抽气量计算公式
            sheet->writeFormula(row, writeColnum + 3, cellCalmin.c_str(), textFormat);
            sheet->writeFormula(row, writeColnum + 4, cellCalmax.c_str(), textFormat);
            // 写入真空暴露时间计算公式
            sheet->writeFormula(row, writeColnum + 5, timeCalSlot.c_str(), textFormat);
            sheet->writeFormula(row, writeColnum + 6, timeCalUnit.c_str(), textFormat);
            sheet->writeFormula(row, writeColnum + 7, timeCalPart.c_str(), textFormat);
            // 写入修正抽气量
            sheet->writeFormula(row, writeColnum + 8, cellFixmin.c_str(), textFormat);
            sheet->writeFormula(row, writeColnum + 9, cellFixmax.c_str(), textFormat);
          }
        }
        else
        {
          if (originData[row - 1][0] > PM0.foilStop && originData[row - 1][0] < PM0.rollBegin \
                                               || originData[row - 1][0] > PM0.rollStop)
          {
            // 长网纸机网部低真空吸水箱、高真空吸水箱、毛毯吸水箱抽气量的TIP计算公式
            unitMinAirflow[row - 1] = originData[row - 1][1] / 1000.0 * originData[row - 1][2] \
                                    * originData[row - 1][3] / 1000.0 * originData[row - 1][4] \
                                    * vacuumMinFactor[vFi][int(originData[row - 1][0]) - 1];
            unitMaxAirflow[row - 1] = originData[row - 1][1] / 1000.0 * originData[row - 1][2] \
                                    * originData[row - 1][3] / 1000.0 * originData[row - 1][4] \
                                    * vacuumMaxFactor[vFi][int(originData[row - 1][0]) - 1];
            // 修正抽气量计算公式
            crtMinFlow[row - 1] = vacFixFlow(referMinVacuum[vFi][int(originData[row - 1][0]) - 1], originData[row - 1][5]) \
                                   * unitMinAirflow[row - 1] * weightFixFlow(stdMaxWeight, maxWeight); // 数据J列
            crtMaxFlow[row - 1] = vacFixFlow(referMaxVacuum[vFi][int(originData[row - 1][0]) - 1], originData[row - 1][6]) \
                                   * unitMaxAirflow[row - 1] * weightFixFlow(stdMinWeight, minWeight); // 数据K列
            // 长网纸机真空暴露时间计算公式
            slotVacTime[row - 1] = originData[row - 1][3] / reelSpeed * 60.0;
            unitVacTime[row - 1] = originData[row - 1][3] * originData[row - 1][2] / reelSpeed * 60.0;
            partVacTime[row - 1] = originData[row - 1][3] * originData[row - 1][2] / reelSpeed * 60.0 \
                                 * originData[row - 1][4];
            // 将公式写入Excel
            if (processControl == 3)
            {
              // 写入标题栏
              sheet->writeStr(0, writeColnum - 1, "refMinVac\n[SI]", titleFormat);
              sheet->writeStr(0, writeColnum    , "refMaxVac\n[SI]", titleFormat);
              sheet->writeStr(0, writeColnum + 1, "minFactor\n[SI]", titleFormat);
              sheet->writeStr(0, writeColnum + 2, "maxFactor\n[SI]", titleFormat);
              sheet->writeStr(0, writeColnum + 3, "airMinFlow\n[m3/min]", titleFormat);
              sheet->writeStr(0, writeColnum + 4, "airMaxFlow\n[m3/min]", titleFormat);
              sheet->writeStr(0, writeColnum + 5, "slotVacTime\n[ms]", titleFormat);
              sheet->writeStr(0, writeColnum + 6, "unitVacTime\n[ms]", titleFormat);
              sheet->writeStr(0, writeColnum + 7, "partVacTime\n[ms]", titleFormat);
              sheet->writeStr(0, writeColnum + 8, "fixMinFlow\n[m3/min]", titleFormat);
              sheet->writeStr(0, writeColnum + 9, "fixMaxFlow\n[m3/min]", titleFormat);
              // 写入TIP参考真空度，最小值和最大值
              sheet->writeNum(row, writeColnum - 1, \
                              referMinVacuum[vFi][int(originData[row - 1][0]) - 1], textFormat);
              sheet->writeNum(row, writeColnum, \
                              referMaxVacuum[vFi][int(originData[row - 1][0]) - 1], textFormat);
              // 写入TIP系数，最小系数和最大系数
              sheet->writeNum(row, writeColnum + 1, \
                              vacuumMinFactor[vFi][int(originData[row - 1][0]) - 1], textFormat);
              sheet->writeNum(row, writeColnum + 2, \
                              vacuumMaxFactor[vFi][int(originData[row - 1][0]) - 1], textFormat);
              // 定义计算使用的单元格行数和列字母
              string rowStr = to_string(row + 1);
              string cellNo_1, cellNo_2, cellNo_3, cellNo_4, cellNo_10, cellNo_11;
              string cellNo_12, cellNo_13;
              // 获取计算使用的第一个单元格编号
              cellNo_1  = transToVocabulary(5)  + rowStr; // F列
              cellNo_2  = transToVocabulary(6)  + rowStr; // G列
              cellNo_3  = transToVocabulary(7)  + rowStr; // H列
              cellNo_4  = transToVocabulary(8)  + rowStr; // I列
              cellNo_10 = transToVocabulary(14) + rowStr; // O列 （最小值）
              cellNo_11 = transToVocabulary(15) + rowStr; // P列 （最大值）
              // 获取修正系数计算使用单元格编号
              cellNo_12 = transToVocabulary(16) + rowStr; // Q列
              cellNo_13 = transToVocabulary(17) + rowStr; // R列
              // 获得抽气量计算公式字符串
              string cellCalmin = cellNo_1 + "/1000*" + cellNo_2 + "*" + cellNo_3 + "/1000*" \
                                + cellNo_4 + "*" + cellNo_10;
              string cellCalmax = cellNo_1 + "/1000*" + cellNo_2 + "*" + cellNo_3 + "/1000*" \
                                + cellNo_4 + "*" + cellNo_11;
              // 获得修正抽气量计算公式字符串
              string cellFixmin = to_string(vacFixFlow(referMinVacuum[vFi][int(originData[row - 1][0]) - 1], originData[row - 1][5])) \
                                + "*" + to_string(weightFixFlow(stdMaxWeight, maxWeight)) + "*" +cellNo_12;
              string cellFixmax = to_string(vacFixFlow(referMinVacuum[vFi][int(originData[row - 1][0]) - 1], originData[row - 1][6])) \
                                + "*" + to_string(weightFixFlow(stdMinWeight, minWeight)) + "*" +cellNo_13;
              // 获得真空暴露时间计算公式字符串
              string timeCalSlot = cellNo_3 + "/" + to_string(reelSpeed) + "*60";
              string timeCalUnit = cellNo_3 + "/" + to_string(reelSpeed) + "*60" + "*" + cellNo_2;
              string timeCalPart = cellNo_3 + "/" + to_string(reelSpeed) + "*60" + "*" + cellNo_2 \
                                 + "*" + cellNo_4;
              // 写入抽气量计算公式
              sheet->writeFormula(row, writeColnum + 3, cellCalmin.c_str(), textFormat);
              sheet->writeFormula(row, writeColnum + 4, cellCalmax.c_str(), textFormat);
              // 写入真空暴露时间计算公式
              sheet->writeFormula(row, writeColnum + 5, timeCalSlot.c_str(), textFormat);
              sheet->writeFormula(row, writeColnum + 6, timeCalUnit.c_str(), textFormat);
              sheet->writeFormula(row, writeColnum + 7, timeCalPart.c_str(), textFormat);
              // 写入修正抽气量
              sheet->writeFormula(row, writeColnum + 8, cellFixmin.c_str(), textFormat);
              sheet->writeFormula(row, writeColnum + 9, cellFixmax.c_str(), textFormat);
            }
          }
          else
          {
            if (originData[row - 1][3] <= 25.4 * 6)
            {
              // 长网纸机网部伏辊、吸移辊、压榨辊抽气量的TIP计算公式（辊径小于6英寸，即152.4mm)
              unitMinAirflow[row - 1] = originData[row - 1][1] / 1000.0 * \
                                        originData[row - 1][3] / 1000.0 * \
                                        vacuumMinFactor[vFi][int(originData[row - 1][0]) - 1] * \
                                        originData[row - 1][4] * originData[row - 1][2];
              unitMaxAirflow[row - 1] = originData[row - 1][1] / 1000.0 * \
                                        originData[row - 1][3] / 1000.0 * \
                                        vacuumMaxFactor[vFi][int(originData[row - 1][0]) - 1] * \
                                        originData[row - 1][4] * originData[row - 1][2];
              // 修正抽气量计算公式
              crtMinFlow[row - 1] = vacFixFlow(referMinVacuum[vFi][int(originData[row - 1][0]) - 1], originData[row - 1][5]) \
                                     * unitMinAirflow[row - 1] * weightFixFlow(stdMaxWeight, maxWeight); // 数据J列
              crtMaxFlow[row - 1] = vacFixFlow(referMaxVacuum[vFi][int(originData[row - 1][0]) - 1], originData[row - 1][6]) \
                                     * unitMaxAirflow[row - 1] * weightFixFlow(stdMinWeight, minWeight); // 数据K列
              // 长网纸机真空暴露时间计算公式
              slotVacTime[row - 1] = originData[row - 1][3] / reelSpeed * 60.0;
              unitVacTime[row - 1] = originData[row - 1][3] * originData[row - 1][2] / reelSpeed * 60.0;
              partVacTime[row - 1] = originData[row - 1][3] * originData[row - 1][2] / reelSpeed * 60.0 \
                                   * originData[row - 1][4];
              // 将公式写入Excel
              if (processControl == 3)
              {
                // 写入标题栏
                sheet->writeStr(0, writeColnum - 1, "refMinVac\n[SI]", titleFormat);
                sheet->writeStr(0, writeColnum    , "refMinVac\n[SI]", titleFormat);
                sheet->writeStr(0, writeColnum + 1, "minFactor\n[SI]", titleFormat);
                sheet->writeStr(0, writeColnum + 2, "maxFactor\n[SI]", titleFormat);
                sheet->writeStr(0, writeColnum + 3, "airMinFlow\n[m3/min]", titleFormat);
                sheet->writeStr(0, writeColnum + 4, "airMaxFlow\n[m3/min]", titleFormat);
                sheet->writeStr(0, writeColnum + 5, "slotVacTime\n[ms]", titleFormat);
                sheet->writeStr(0, writeColnum + 6, "unitVacTime\n[ms]", titleFormat);
                sheet->writeStr(0, writeColnum + 7, "partVacTime\n[ms]", titleFormat);
                sheet->writeStr(0, writeColnum + 8, "fixMinFlow\n[m3/min]", titleFormat);
                sheet->writeStr(0, writeColnum + 9, "fixMaxFlow\n[m3/min]", titleFormat);
                // 写入TIP参考真空度，最小值和最大值
                sheet->writeNum(row, writeColnum - 1, \
                                referMinVacuum[vFi][int(originData[row - 1][0]) - 1], textFormat);
                sheet->writeNum(row, writeColnum, \
                                referMaxVacuum[vFi][int(originData[row - 1][0]) - 1], textFormat);
                // 写入TIP系数，最小系数和最大系数
                sheet->writeNum(row, writeColnum + 1, \
                                vacuumMinFactor[vFi][int(originData[row - 1][0]) - 1], textFormat);
                sheet->writeNum(row, writeColnum + 2, \
                                vacuumMaxFactor[vFi][int(originData[row - 1][0]) - 1], textFormat);
                // 定义计算使用的单元格行数和列字母
                string rowStr = to_string(row + 1);
                string cellNo_1, cellNo_2, cellNo_3, cellNo_4, cellNo_10, cellNo_11;
                string cellNo_12, cellNo_13;
                // 获取计算使用的单元格编号
                cellNo_1  = transToVocabulary(5)  + rowStr; // F列
                cellNo_2  = transToVocabulary(6)  + rowStr; // G列
                cellNo_3  = transToVocabulary(7)  + rowStr; // H列
                cellNo_4  = transToVocabulary(8)  + rowStr; // I列
                cellNo_10 = transToVocabulary(14) + rowStr; // O列 （最小值）
                cellNo_11 = transToVocabulary(15) + rowStr; // P列 （最大值）
                // 获取修正系数计算使用单元格编号
                cellNo_12 = transToVocabulary(16) + rowStr; // Q列
                cellNo_13 = transToVocabulary(17) + rowStr; // R列
                // 获得抽气量计算公式字符串
                string cellCalmin = cellNo_1 + "/1000*" + cellNo_3 + "/1000*" + cellNo_10 + "*" + cellNo_4 \
                                  + "*" + cellNo_2;
                string cellCalmax = cellNo_1 + "/1000*" + cellNo_3 + "/1000*" + cellNo_11 + "*" + cellNo_4 \
                                  + "*" + cellNo_2;
                // 获得修正抽气量计算公式字符串
                string cellFixmin = to_string(vacFixFlow(referMinVacuum[vFi][int(originData[row - 1][0]) - 1],\
                                    originData[row - 1][5])) \
                                  + "*" + to_string(weightFixFlow(stdMaxWeight, maxWeight)) + "*" +cellNo_12;
                string cellFixmax = to_string(vacFixFlow(referMinVacuum[vFi][int(originData[row - 1][0]) - 1],\
                                    originData[row - 1][6])) \
                                  + "*" + to_string(weightFixFlow(stdMinWeight, minWeight)) + "*" +cellNo_13;
                // 获得真空暴露时间计算公式字符串
                string timeCalSlot = cellNo_3 + "/" + to_string(reelSpeed) + "*60";
                string timeCalUnit = cellNo_3 + "/" + to_string(reelSpeed) + "*60" + "*" + cellNo_2;
                string timeCalPart = cellNo_3 + "/" + to_string(reelSpeed) + "*60" + "*" + cellNo_2 \
                                   + "*" + cellNo_4;
                // 写入抽气量计算公式
                sheet->writeFormula(row, writeColnum + 3, cellCalmin.c_str(), textFormat);
                sheet->writeFormula(row, writeColnum + 4, cellCalmax.c_str(), textFormat);
                // 写入真空暴露时间计算公式
                sheet->writeFormula(row, writeColnum + 5, timeCalSlot.c_str(), textFormat);
                sheet->writeFormula(row, writeColnum + 6, timeCalUnit.c_str(), textFormat);
                sheet->writeFormula(row, writeColnum + 7, timeCalPart.c_str(), textFormat);
                // 写入修正抽气量
                sheet->writeFormula(row, writeColnum + 8, cellFixmin.c_str(), textFormat);
                sheet->writeFormula(row, writeColnum + 9, cellFixmax.c_str(), textFormat);
              }
            }
            else
            {
              double arcLength = chord2arcLength(originData[row - 1][7], originData[row -1][3]);
              // 长网纸机网部伏辊、吸移辊、压榨辊抽气量的TIP计算公式（辊径大于6英寸，即152.4mm)
              unitMinAirflow[row - 1] = originData[row - 1][1] / 1000.0 * arcLength / 1000.0 \
                                      * vacuumMinFactor[vFi][int(originData[row - 1][0]) - 1] \
                                      * originData[row - 1][2];
              unitMaxAirflow[row - 1] = originData[row - 1][1] / 1000.0 * arcLength / 1000.0 \
                                      * vacuumMaxFactor[vFi][int(originData[row - 1][0]) - 1] \
                                      * originData[row - 1][2];
              // 修正抽气量计算公式
              crtMinFlow[row - 1] = vacFixFlow(referMinVacuum[vFi][int(originData[row - 1][0]) - 1], originData[row - 1][5]) \
                                     * unitMinAirflow[row - 1] * weightFixFlow(stdMaxWeight, maxWeight); // 数据J列
              crtMaxFlow[row - 1] = vacFixFlow(referMaxVacuum[vFi][int(originData[row - 1][0]) - 1], originData[row - 1][6]) \
                                     * unitMaxAirflow[row - 1] * weightFixFlow(stdMinWeight, minWeight); // 数据K列
              // 长网纸机真空暴露时间计算公式
              slotVacTime[row - 1] = originData[row - 1][3] / reelSpeed * 60.0;
              unitVacTime[row - 1] = originData[row - 1][3] * originData[row - 1][2] / reelSpeed * 60.0;
              partVacTime[row - 1] = originData[row - 1][3] * originData[row - 1][2] / reelSpeed * 60.0 \
                                   * originData[row - 1][4];
              // 将公式写入Excel
              if (processControl == 3)
              {
                // 写入标题栏
                sheet->writeStr(0, writeColnum - 1, "refMinVac\n[SI]", titleFormat);
                sheet->writeStr(0, writeColnum    , "refMinVac\n[SI]", titleFormat);
                sheet->writeStr(0, writeColnum + 1, "MinFactor\n[SI]", titleFormat);
                sheet->writeStr(0, writeColnum + 2, "MaxFactor\n[SI]", titleFormat);
                sheet->writeStr(0, writeColnum + 3, "airMinFlow\n[m3/min]", titleFormat);
                sheet->writeStr(0, writeColnum + 4, "airMaxFlow\n[m3/min]", titleFormat);
                sheet->writeStr(0, writeColnum + 5, "slotVacTime\n[ms]", titleFormat);
                sheet->writeStr(0, writeColnum + 6, "unitVacTime\n[ms]", titleFormat);
                sheet->writeStr(0, writeColnum + 7, "partVacTime\n[ms]", titleFormat);
                sheet->writeStr(0, writeColnum + 8, "fixMinFlow\n[m3/min]", titleFormat);
                sheet->writeStr(0, writeColnum + 9, "fixMaxFlow\n[m3/min]", titleFormat);
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
                string rowStr = to_string(row + 1);
                string cellNo_1, cellNo_2, cellNo_3, cellNo_4, cellNo_7, cellNo_10, cellNo_11;
                string cellNo_12, cellNo_13;
                // 获取计算使用的单元格编号
                cellNo_1  = transToVocabulary(5)  + rowStr; // F列
                cellNo_2  = transToVocabulary(6)  + rowStr; // G列
                cellNo_3  = transToVocabulary(7)  + rowStr; // H列
                cellNo_4  = transToVocabulary(8)  + rowStr; // I列
                cellNo_7  = transToVocabulary(11) + rowStr; // L列
                cellNo_10 = transToVocabulary(14) + rowStr; // O列 （最小值）
                cellNo_11 = transToVocabulary(15) + rowStr; // P列 （最大值）
                // 获取修正系数计算使用单元格编号
                cellNo_12 = transToVocabulary(16) + rowStr; // Q列
                cellNo_13 = transToVocabulary(17) + rowStr; // R列
                // 获得计算公式字符串
                string cellCalmin = cellNo_1 + "/1000*" + cellNo_7 + "/2*asin(1-0.5*power(" \
                                  + cellNo_3 + "/" + cellNo_7 + "*2,2))" + "/1000*" + cellNo_10 \
                                  + "*" + cellNo_4 + "*" + cellNo_2;
                string cellCalmax = cellNo_1 + "/1000*" + cellNo_7 + "/2*asin(1-0.5*power(" \
                                  + cellNo_3 + "/" + cellNo_7 + "*2,2))" + "/1000*" + cellNo_11 \
                                  + "*" + cellNo_4 + "*" + cellNo_2;
                // 获得修正抽气量计算公式字符串
                string cellFixmin = to_string(vacFixFlow(referMinVacuum[vFi][int(originData[row - 1][0]) - 1], \
                                    originData[row - 1][5])) \
                                  + "*" + to_string(weightFixFlow(stdMaxWeight, maxWeight)) + "*" +cellNo_12;
                string cellFixmax = to_string(vacFixFlow(referMaxVacuum[vFi][int(originData[row - 1][0]) - 1], \
                                    originData[row - 1][6])) \
                                  + "*" + to_string(weightFixFlow(stdMinWeight, minWeight)) + "*" +cellNo_13;
                // 获得真空暴露时间计算公式字符串
                string timeCalSlot = cellNo_3 + "/" + to_string(reelSpeed) + "*60";
                string timeCalUnit = cellNo_3 + "/" + to_string(reelSpeed) + "*60" + "*" + cellNo_2;
                string timeCalPart = cellNo_3 + "/" + to_string(reelSpeed) + "*60" + "*" + cellNo_2 \
                                   + "*" + cellNo_4;
                // 写入抽气量计算公式，c_str()将字符串（string）转换为（char）
                sheet->writeFormula(row, writeColnum + 3, cellCalmin.c_str(), textFormat);
                sheet->writeFormula(row, writeColnum + 4, cellCalmax.c_str(), textFormat);
                // 写入真空暴露时间计算公式
                sheet->writeFormula(row, writeColnum + 5, timeCalSlot.c_str(), textFormat);
                sheet->writeFormula(row, writeColnum + 6, timeCalUnit.c_str(), textFormat);
                sheet->writeFormula(row, writeColnum + 7, timeCalPart.c_str(), textFormat);
                // 写入修正抽气量
                sheet->writeFormula(row, writeColnum + 8, cellFixmin.c_str(), textFormat);
                sheet->writeFormula(row, writeColnum + 9, cellFixmax.c_str(), textFormat);
              }
            }
          }
        }
        // 计算总抽气量
        totalMinAirflow = totalMinAirflow + unitMinAirflow[row - 1];
        totalMaxAirflow = totalMaxAirflow + unitMaxAirflow[row - 1];
        // 计算修正总抽气量
        totalCrtMinFlow = totalCrtMinFlow + crtMinFlow[row - 1];
        totalCrtMaxFlow = totalCrtMaxFlow + crtMaxFlow[row - 1];
        // 输出不同抽吸位置计算的抽气量至屏幕
        std::cout << suctionLocation[row - 1] << "\t"
             << formatString(to_string(unitMinAirflow[row - 1]), " ", 12)
             << formatString(to_string(unitMaxAirflow[row - 1]), " ", 12)
             << formatString(to_string(slotVacTime[row - 1]), " ", 12)
             << formatString(to_string(unitVacTime[row - 1]), " ", 12)
             << formatString(to_string(partVacTime[row - 1]), " ", 12)
             << formatString(to_string(crtMinFlow[row - 1]), " ", 12)
             << formatString(to_string(crtMaxFlow[row - 1]), " ", 12)
             << endl;
        // 将计算数据结果写入Excel
        if (processControl == 2)
        {
          // 写入标题
          sheet->writeStr(0, writeColnum - 1, "refMinVac\n[SI]", titleFormat);
          sheet->writeStr(0, writeColnum    , "refMinVac\n[SI]", titleFormat);
          sheet->writeStr(0, writeColnum + 1, "minFactor\n[SI]", titleFormat);
          sheet->writeStr(0, writeColnum + 2, "maxFactor\n[SI]", titleFormat);
          sheet->writeStr(0, writeColnum + 3, "airMinFlow\n[m3/min]", titleFormat);
          sheet->writeStr(0, writeColnum + 4, "airMaxFlow\n[m3/min]", titleFormat);
          sheet->writeStr(0, writeColnum + 5, "slotVacTime\n[ms]", titleFormat);
          sheet->writeStr(0, writeColnum + 6, "unitVacTime\n[ms]", titleFormat);
          sheet->writeStr(0, writeColnum + 7, "partVacTime\n[ms]", titleFormat);
          sheet->writeStr(0, writeColnum + 8, "fixMinFlow\n[m3/min]", titleFormat);
          sheet->writeStr(0, writeColnum + 9, "fixMaxFlow\n[m3/min]", titleFormat);
          // 写入计算结果至Excel
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
          sheet->writeNum(row, writeColnum + 5, slotVacTime[row - 1], textFormat);
          sheet->writeNum(row, writeColnum + 6, unitVacTime[row - 1], textFormat);
          sheet->writeNum(row, writeColnum + 7, partVacTime[row - 1], textFormat);
          sheet->writeNum(row, writeColnum + 8, crtMinFlow[row - 1], textFormat);
          sheet->writeNum(row, writeColnum + 9, crtMaxFlow[row - 1], textFormat);
        }
      }
      // 输出总抽气量到Excel
      if (processControl == 2)
      {
        // 判断原文件是否有统计行，若有则不写入下一行
        if (!conCal)
        {
          sheet->writeStr(lastRow, writeColnum + 2, "Total", textFormat);
          sheet->writeNum(lastRow, writeColnum + 3, totalMinAirflow, textFormat);
          sheet->writeNum(lastRow, writeColnum + 4, totalMaxAirflow, textFormat);
          sheet->writeNum(lastRow, writeColnum + 8, totalCrtMinFlow, textFormat);
          sheet->writeNum(lastRow, writeColnum + 9, totalCrtMaxFlow, textFormat);
        }
        // 若原文件没有统计行，则写入下一行
        else
        {
          sheet->writeNum(lastRow, writeColnum + 3, totalMinAirflow, textFormat);
          sheet->writeNum(lastRow, writeColnum + 4, totalMaxAirflow, textFormat);
        }
      }
      if (processControl == 3)
      {
        // 判断原文件是否有统计行，若有则不写入下一行
        if (!conCal)
        {
          sheet->writeStr(lastRow, writeColnum + 2, "Total", textFormat);
          // 列出总最小抽气量和总最大抽气量的公式
          string minAirForm = "SUM(" + transToVocabulary(writeColnum + 3) + to_string(2) + ":"\
                            + transToVocabulary(writeColnum + 3) + to_string(lastRow) + ")";
          string maxAirForm = "SUM(" + transToVocabulary(writeColnum + 4) + to_string(2) + ":"\
                            + transToVocabulary(writeColnum + 4) + to_string(lastRow) + ")";
          // 列出修正总最小抽气量和总最大抽气量的公式
          string crtMinAirForm = "SUM(" + transToVocabulary(writeColnum + 8) + to_string(2) + ":"\
                            + transToVocabulary(writeColnum + 8) + to_string(lastRow) + ")";
          string crtMaxAirForm = "SUM(" + transToVocabulary(writeColnum + 9) + to_string(2) + ":"\
                            + transToVocabulary(writeColnum + 9) + to_string(lastRow) + ")";
          sheet->writeFormula(lastRow, writeColnum + 3, minAirForm.c_str(), textFormat);
          sheet->writeFormula(lastRow, writeColnum + 4, maxAirForm.c_str(), textFormat);
          sheet->writeFormula(lastRow, writeColnum + 8, crtMinAirForm.c_str(), textFormat);
          sheet->writeFormula(lastRow, writeColnum + 9, crtMaxAirForm.c_str(), textFormat);
        }
        // 若原文件没有统计行，则写入下一行
        else
        {
          // 列出总最小抽气量和总最大抽气量的公式
          string minAirForm = "SUM(" + transToVocabulary(writeColnum + 3) + to_string(2) + ":"\
                            + transToVocabulary(writeColnum + 3) + to_string(lastRow) + ")";
          string maxAirForm = "SUM(" + transToVocabulary(writeColnum + 4) + to_string(2) + ":"\
                            + transToVocabulary(writeColnum + 4) + to_string(lastRow) + ")";
          // 列出修正总最小抽气量和总最大抽气量的公式
          string crtMinAirForm = "SUM(" + transToVocabulary(writeColnum + 8) + to_string(2) + ":"\
                            + transToVocabulary(writeColnum + 8) + to_string(lastRow) + ")";
          string crtMaxAirForm = "SUM(" + transToVocabulary(writeColnum + 9) + to_string(2) + ":"\
                            + transToVocabulary(writeColnum + 9) + to_string(lastRow) + ")";
          sheet->writeFormula(lastRow, writeColnum + 3, minAirForm.c_str(), textFormat);
          sheet->writeFormula(lastRow, writeColnum + 4, maxAirForm.c_str(), textFormat);
          sheet->writeFormula(lastRow, writeColnum + 8, crtMinAirForm.c_str(), textFormat);
          sheet->writeFormula(lastRow, writeColnum + 9, crtMaxAirForm.c_str(), textFormat);
        }
      }

      // 输出总抽气量到屏幕
      std::cout << formatString("------Total Vacuum Airflow is: ", "-") << "\t"
           << formatString(to_string(totalMinAirflow), " ", 12)
           << formatString(to_string(totalMaxAirflow), " ", 12)
           << formatString("-", "-", 12)
           << formatString("-", "-", 12)
           << formatString("-", "-", 12)
           << formatString(to_string(totalCrtMinFlow), " ", 12)
           << formatString(to_string(totalCrtMaxFlow), " ", 12)
           << endl;
//      lastRowVal = "NULL";
    }
  }
  else
  {
    // 装载数据表格错误，输出错误信息
    std::cout << bookRead->errorMessage() << endl;
  }

  // 储存数据表格
  if (!bookRead->save("originData.xlsx"))
  {
    // 存储失败输出错误信息
    std::cout << bookRead->errorMessage() << endl;
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
  std::cout << "\nPrint any key to continue..." << endl;
  _getch();
  return 0;
}
