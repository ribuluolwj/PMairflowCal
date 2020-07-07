#include <cmath>
#include <iostream>
#include <conio.h>
#include "rpfunc.hpp"

using namespace std;
using namespace basicmath;

//double chord2ArcLength(double diameter, double chord)
//{
//  double r = diameter / 2.0;
//  double arcLength = r * acos(1.0 - 0.5 * pow(chord / r, 2));
//  return arcLength;
//}
//
//double arcAngle2Chord(double diameter, double arcAngleD)
//{
//  double arcAngleR = arcAngleD / 90.0 * asin(1.0);
//  double chord = sqrt(2.0 * pow(diameter / 2.0, 2) * (1.0 - cos(arcAngleR)));
//  return chord;
//}

int main()
{
  cout << "Choose calculation do you want:" << "\n"
       << "[1] Calculate chord;" << "\n"
       << "[2] Calculate arc length." << endl;
     // 判断读取数据是否正确
  int process;
  while (1)
  {
    // 定义临时字符数组用来存取屏幕读入的数据
    char calType[6] = "1";
    cin.getline(calType, 6);
    // 将读取的输入字符数组转换为字符串
    string pT = calType;
    // 定义输入数据正确性判定字符串
    string ans = "1,2";
    string::size_type idx;
    // 查找输入数据是否在判定字符串中
    idx = ans.find(pT);
    // 判定输入是否正确
    if (idx == string::npos || pT.length() > 1)
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
        process = stoi(pT);
        break;
      }
    }
  }

  if (process == 1)
  {
    double diameterNum, arcAngleNum;
    double chord;
    cout << "Input diameter:" << endl;
    cin >> diameterNum;
    cout << "Input arc angle:" << endl;
    cin >> arcAngleNum;
    chord = angle2Chord(diameterNum, arcAngleNum);
    cout << "Chord is: "
         << "\t" << chord << endl;
  }
  else
  {
    if (process == 2)
    {
      double diameterNum, chordNum;
      double arcLength;
      cout << "Input diameter:" << endl;
      cin >> diameterNum;
      cout << "Input chord length:" << endl;
      cin >> chordNum;
      arcLength = chord2arcLength(diameterNum, chordNum);
      cout << "Arc length is:"
           << "\t" << arcLength << endl;
    }
  }
  _getch();
  return 0;
}
