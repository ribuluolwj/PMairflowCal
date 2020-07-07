#pragma once
//#ifndef RPFUNC_HPP_
//#define RPFUNC_HPP_
#include <iostream>
#include <cmath>
#include <GL/gl.h>

using namespace std;

// 定义基本数学计算工具命名空间
namespace basicmath
{
  // 定义根据弦长求弧长函数
  double chord2arcLength(double diameter, double chord)
  {
    double r = diameter / 2.0;
    double arcLength = r * acos(1.0 - 0.5 * pow(chord / r, 2));
    return arcLength;
  }
  // 定义根据角度求弦长函数
  double angle2Chord(double diameter, double arcAngle, bool rad = false)
  {
    // 默认使用角度制
    // 判断是否使用角度制
    if (!rad)
    {
      // 将角度制换成弧度制
      double arcAngleR = arcAngle / 90.0 * asin(1.0);
    }
    else
    {
      double anrAngleR = arcAngle;
    }
    // 根据角度求弦长
    double chord = sqrt(2.0 * pow(diameter / 2.0, 2) * (1.0 - cos(arcAngle)));
    return chord;
  }

  // 定义根据角度求弧长函数
  double angle2arcLength(double diameter, double arcAngle, bool rad = false)
  {
    // 默认使用角度制
    // 判断是否使用角度制
    if (!rad)
    {
      // 将角度制换成弧度制
      double arcAngleR = arcAngle / 90.0 * asin(1.0);
    }
    else
    {
      double anrAngleR = arcAngle;
    }
    double arcLength= arcAngle * diameter / 2.0;
    return arcLength;
  }

  // 定义真空度体积流量修正函数
  double vacFixFlow(double stdVac, double realVac)
  {
    double vacFix;
    vacFix = sqrt(realVac / stdVac);
//    std::cout << stdVac << "\t" << realVac << "\t" << vacFix << endl;

    return vacFix;
  }

  // 定义真空度质量流量修正函数
  double vacFixMassFlow(double stdVac, double realVac, double atm = 101.325)
  {
    double vacFix;
    vacFix = sqrt(realVac / stdVac) * sqrt((atm - realVac) / (atm - stdVac));
    return vacFix;
  }

  // 定义定量气体流量修正函数
  double weightFixFlow(double stdWeight, double realWeight, std::string module = "sample")
  {
    double weightFix;
    if (module == "no")
    {
      weightFix = 1.0;
    }
    else
    {
      if (module == "sample")
      {
        weightFix = sqrt(stdWeight / realWeight);
      }
    }
    return weightFix;
  }

  // 定义定量质量流量修正函数
  double weightFixMassFlow(double stdWeight, double realWeight, double atm = 101.325)
  {
    double weightFix;

    return weightFix;
  }

//  // 定义TIPCal程序使用的绘图函数
//  void drawPaperMachine(std::string pmType)
//  {
//
//  }
}; // namespace basicmath

namespace basicformat
{
  // 格式化输出字符串，左对齐，参数：字符串，填充类型(默任“空格”），字符串长度（默认60）
  std::string formatString(std::string str, std::string slot = " ", int length = 60)
  {
    std::string sb;
    sb.append(str);
    int count = length - str.length();
    while(count > 0){
        sb.append(slot);
        count --;
    }
    return sb;
  }
  // 将数字转化为对应顺序的英文字母，默认大写，若为小写，第二个变量写为'S'
  std::string transToVocabulary(int number, const char bigOrSmall = 'B')
  {
    // 定义：字母对应的字符常量
    char vocabulary;
    // 定义：中转用的字符数组
    char vb[2];
    // 定义：字母字符串
    std::string vob;
    // 判断转换成大写或者小写
    switch (bigOrSmall)
    {
    case 'S':
      vocabulary = 97 + number;
      vb[0] = vocabulary;
      vob = vb[0];
      break;
    default:
      vocabulary = 65 + number;
      vb[0] = vocabulary;
      vob = vb[0];
      break;
    }
    return vob;
  }

} // namespace basicformat

namespace calconfig
{
  // 定义纸机类，用来判断真空元件类型分界线
  class Papermachine
  {
    public:
    // 定义个真空元件分类界限变量
      int foilStop, flatStop, rollBegin, rollStop, feltBegin;
    // 定义设置真空元件分类界限变量函数
      void setVacuumConfig(int foilboxS, int flatboxS, int rollB, int rollS, int feltboxB)
      {
        foilStop  = foilboxS;
        flatStop  = flatboxS;
        rollBegin    = rollB;
        rollStop     = rollS;
        feltBegin = feltboxB;
      }
    // PM0.foilStop......吸湿箱最大工段号，1~3全部使用粗算公式，并不考虑实际暴露面积，至考虑元件数量，真空度低，无需校准
    // PM0.flatStop......真空箱最大工段号，4~5，分别为低真空和高真空吸水箱，4之后是水线，此时计算公式考虑暴露面积
    // PM0.rollBegin.....真空吸辊开始工段号，6~21，分别为伏辊低真空、高真空，吸移辊低、高真空，压榨辊低真空、高真空
    // PM0.rollStop......真空吸辊结束工段号
    // PM0.feltBegin.....毛毯吸水箱和转移箱开始工段号，22~25
  };
  // 定义输出程序开头文字声明
  void printHeadClaim()
  {
    cout << "[1 ] 瓦楞芯纸      （CORRUGATING MEDIUM）；\n"
         << "[2 ] 复印纸和书写纸（PRINTING & WRITING）；\n"
         << "[3 ] 书本纸        （BOOK PAPERS）；\n"
         << "[4 ] 高定量印刷纸  （HEAVYWEIGHTS）；\n"
         << "[5 ] MG & MF       （MG & MF PAPERS）；\n"
         << "[6 ] 格拉辛纸      （GLASSINE, GREASEPROOF）；\n"
         << "[7 ] 复写原纸      （CARBONIZING）；\n"
         << "[8 ] 蜡基特纸      （WAXING BASE）；\n"
         << "[9 ] 卷烟纸        （CIGARETTE, CONDENSOR TISSUE）；\n"
         << "[10] 餐巾纸        （NAPKIN）；\n"
         << "[11] 双层毛巾纸    （TOWEL, TWO PLY）；\n"
         << "[12] 单层毛巾纸    （TOWEL, SINGLE PLY）；\n"
         << "[13] 目录纸        （DIRECTORY ROTO, CATALOG）；\n"
         << "[14] SC杂志纸      （SC MAGAZINE）；\n"
         << "[15] LWC出版纸     （LWC PUBLICATION）；\n"
         << "[16] 新闻纸        （NEWSPRINT）；\n"
         << "[17] 纸袋纸        （BAG）；\n"
         << "[18] 饱和纸        （SATURATING）；\n"
         << "[19] 低定量线板纸  （LINERBOARD-LW）；\n"
         << "[20] 高定量线板纸  （LINERBOARD-HW）；\n"
         << "[21] 白芯白卡纸    （SBS PAPERBOARD）；\n"
         << "[22] 特种包装纸    （PACKAGING SPECIALTIES）；\n"
         << "[23] 浆板          （PULP）。" << endl;
    cout << "请输入计算纸种参数（若同时计算多种纸种，或原始数据表中已选定，可跳过输入）：";
  }
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
} // namespace papermachine
//#endif
