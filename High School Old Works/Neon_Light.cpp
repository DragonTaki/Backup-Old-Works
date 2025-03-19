/*----- ----- ----- ----- ----- -----*/
//             Neon light            //
//                                   //
//Auther : Synasaivaltos Terdolic    //
//Company:                           //
//Date   : 2014/12/31                //
/*----- ----- ----- ----- ----- -----*/

#include <cstdlib>
#include <ctime>
#include <iostream>
#include <windows.h>

#define time_limit 500

using namespace std;

void basic_frame(void);
void change_color(int color);

int main(void)
{
	srand(int(time(0)));
	
	int color;
	int lines;
	int time;
	int temp;
	
	basic_frame();
	Sleep(5000);
	system("cls");
	
	color=rand()%240;
	change_color(color);
	for(int i=0;i<920;i++)
	{
		temp=rand()%2;
		if(temp==0)
		{
			cout << "ー";
		}
		else
		{
			cout << "―";
		}
	}
	
	while(1)
	{
		color=rand()%240;
		change_color(color);
		lines=rand()%10+1;
		for(int i=0;i<lines;i++)
		{
			for(int j=0;j<40;j++)
			{
				temp=rand()%2;
				if(temp==0)
				{
					cout << "ー";
				}
				else
				{
					cout << "―";
				}
			}
		}
		
		//system("pause");
		time=rand()%time_limit+1;
		Sleep(time);
		//system("cls");
	}
	
	system("pause");
	return 0;
}

void basic_frame(void)
{
	system("cls");
	cout << "----- ----- ----- ----- ----- ----- ----- ----- ----- ----- ----- -----" << endl;
	cout << "|                              Neon light                             |" << endl;
	cout << "----- ----- ----- ----- ----- ----- ----- ----- ----- ----- ----- -----" << endl;
	cout << "|                    Made by Synasaivaltos Terdolic                   |" << endl;
	cout << "|                       synasaivaltos@gmail.com                       |" << endl;
	cout << "|                               Ver. 1.1                              |" << endl;
	cout << "----- ----- ----- ----- ----- ----- ----- ----- ----- ----- ----- -----" << endl;
	
	return;
}

void change_color(int color)
{
	switch(color)
		{
			case 0:
				system("color 01");
				break;
			case 1:
				system("color 02");
				break;
			case 2:
				system("color 03");
				break;
			case 3:
				system("color 04");
				break;
			case 4:
				system("color 05");
				break;
			case 5:
				system("color 06");
				break;
			case 6:
				system("color 07");
				break;
			case 7:
				system("color 08");
				break;
			case 8:
				system("color 09");
				break;
			case 9:
				system("color 0A");
				break;
			case 10:
				system("color 0B");
				break;
			case 11:
				system("color 0C");
				break;
			case 12:
				system("color 0D");
				break;
			case 13:
				system("color 0E");
				break;
			case 14:
				system("color 0F");
				break;
			case 15:
				system("color 10");
				break;
			case 16:
				system("color 12");
				break;
			case 17:
				system("color 13");
				break;
			case 18:
				system("color 14");
				break;
			case 19:
				system("color 15");
				break;
			case 20:
				system("color 16");
				break;
			case 21:
				system("color 17");
				break;
			case 22:
				system("color 18");
				break;
			case 23:
				system("color 19");
				break;
			case 24:
				system("color 1A");
				break;
			case 25:
				system("color 1B");
				break;
			case 26:
				system("color 1C");
				break;
			case 27:
				system("color 1D");
				break;
			case 28:
				system("color 1E");
				break;
			case 29:
				system("color 1F");
				break;
			case 30:
				system("color 20");
				break;
			case 31:
				system("color 21");
				break;
			case 32:
				system("color 23");
				break;
			case 33:
				system("color 24");
				break;
			case 34:
				system("color 25");
				break;
			case 35:
				system("color 26");
				break;
			case 36:
				system("color 27");
				break;
			case 37:
				system("color 28");
				break;
			case 38:
				system("color 29");
				break;
			case 39:
				system("color 2A");
				break;
			case 40:
				system("color 2B");
				break;
			case 41:
				system("color 2C");
				break;
			case 42:
				system("color 2D");
				break;
			case 43:
				system("color 2E");
				break;
			case 44:
				system("color 2F");
				break;
			case 45:
				system("color 30");
				break;
			case 46:
				system("color 31");
				break;
			case 47:
				system("color 32");
				break;
			case 48:
				system("color 34");
				break;
			case 49:
				system("color 35");
				break;
			case 50:
				system("color 36");
				break;
			case 51:
				system("color 37");
				break;
			case 52:
				system("color 38");
				break;
			case 53:
				system("color 39");
				break;
			case 54:
				system("color 3A");
				break;
			case 55:
				system("color 3B");
				break;
			case 56:
				system("color 3C");
				break;
			case 57:
				system("color 3D");
				break;
			case 58:
				system("color 3E");
				break;
			case 59:
				system("color 3F");
				break;
			case 60:
				system("color 40");
				break;
			case 61:
				system("color 41");
				break;
			case 62:
				system("color 42");
				break;
			case 63:
				system("color 43");
				break;
			case 64:
				system("color 45");
				break;
			case 65:
				system("color 46");
				break;
			case 66:
				system("color 47");
				break;
			case 67:
				system("color 48");
				break;
			case 68:
				system("color 49");
				break;
			case 69:
				system("color 4A");
				break;
			case 70:
				system("color 4B");
				break;
			case 71:
				system("color 4C");
				break;
			case 72:
				system("color 4D");
				break;
			case 73:
				system("color 4E");
				break;
			case 74:
				system("color 4F");
				break;
			case 75:
				system("color 50");
				break;
			case 76:
				system("color 51");
				break;
			case 77:
				system("color 52");
				break;
			case 78:
				system("color 53");
				break;
			case 79:
				system("color 54");
				break;
			case 80:
				system("color 56");
				break;
			case 81:
				system("color 57");
				break;
			case 82:
				system("color 58");
				break;
			case 83:
				system("color 59");
				break;
			case 84:
				system("color 5A");
				break;
			case 85:
				system("color 5B");
				break;
			case 86:
				system("color 5C");
				break;
			case 87:
				system("color 5D");
				break;
			case 88:
				system("color 5E");
				break;
			case 89:
				system("color 5F");
				break;
			case 90:
				system("color 60");
				break;
			case 91:
				system("color 61");
				break;
			case 92:
				system("color 62");
				break;
			case 93:
				system("color 63");
				break;
			case 94:
				system("color 64");
				break;
			case 95:
				system("color 65");
				break;
			case 96:
				system("color 67");
				break;
			case 97:
				system("color 68");
				break;
			case 98:
				system("color 69");
				break;
			case 99:
				system("color 6A");
				break;
			case 100:
				system("color 6B");
				break;
			case 101:
				system("color 6C");
				break;
			case 102:
				system("color 6D");
				break;
			case 103:
				system("color 6E");
				break;
			case 104:
				system("color 6F");
				break;
			case 105:
				system("color 70");
				break;
			case 106:
				system("color 71");
				break;
			case 107:
				system("color 72");
				break;
			case 108:
				system("color 73");
				break;
			case 109:
				system("color 74");
				break;
			case 110:
				system("color 75");
				break;
			case 111:
				system("color 76");
				break;
			case 112:
				system("color 78");
				break;
			case 113:
				system("color 79");
				break;
			case 114:
				system("color 7A");
				break;
			case 115:
				system("color 7B");
				break;
			case 116:
				system("color 7C");
				break;
			case 117:
				system("color 7D");
				break;
			case 118:
				system("color 7E");
				break;
			case 119:
				system("color 7F");
				break;
			case 120:
				system("color 80");
				break;
			case 121:
				system("color 81");
				break;
			case 122:
				system("color 82");
				break;
			case 123:
				system("color 83");
				break;
			case 124:
				system("color 84");
				break;
			case 125:
				system("color 85");
				break;
			case 126:
				system("color 86");
				break;
			case 127:
				system("color 87");
				break;
			case 128:
				system("color 89");
				break;
			case 129:
				system("color 8A");
				break;
			case 130:
				system("color 8B");
				break;
			case 131:
				system("color 8C");
				break;
			case 132:
				system("color 8D");
				break;
			case 133:
				system("color 8E");
				break;
			case 134:
				system("color 8F");
				break;
			case 135:
				system("color 90");
				break;
			case 136:
				system("color 91");
				break;
			case 137:
				system("color 92");
				break;
			case 138:
				system("color 93");
				break;
			case 139:
				system("color 94");
				break;
			case 140:
				system("color 95");
				break;
			case 141:
				system("color 96");
				break;
			case 142:
				system("color 97");
				break;
			case 143:
				system("color 98");
				break;
			case 144:
				system("color 9A");
				break;
			case 145:
				system("color 9B");
				break;
			case 146:
				system("color 9C");
				break;
			case 147:
				system("color 9D");
				break;
			case 148:
				system("color 9E");
				break;
			case 149:
				system("color 9F");
				break;
			case 150:
				system("color A0");
				break;
			case 151:
				system("color A1");
				break;
			case 152:
				system("color A2");
				break;
			case 153:
				system("color A3");
				break;
			case 154:
				system("color A4");
				break;
			case 155:
				system("color A5");
				break;
			case 156:
				system("color A6");
				break;
			case 157:
				system("color A7");
				break;
			case 158:
				system("color A8");
				break;
			case 159:
				system("color A9");
				break;
			case 160:
				system("color AB");
				break;
			case 161:
				system("color AC");
				break;
			case 162:
				system("color AD");
				break;
			case 163:
				system("color AE");
				break;
			case 164:
				system("color AF");
				break;
			case 165:
				system("color B0");
				break;
			case 166:
				system("color B1");
				break;
			case 167:
				system("color B2");
				break;
			case 168:
				system("color B3");
				break;
			case 169:
				system("color B4");
				break;
			case 170:
				system("color B5");
				break;
			case 171:
				system("color B6");
				break;
			case 172:
				system("color B7");
				break;
			case 173:
				system("color B8");
				break;
			case 174:
				system("color B9");
				break;
			case 175:
				system("color BA");
				break;
			case 176:
				system("color BC");
				break;
			case 177:
				system("color BD");
				break;
			case 178:
				system("color BE");
				break;
			case 179:
				system("color BF");
				break;
			case 180:
				system("color C0");
				break;
			case 181:
				system("color C1");
				break;
			case 182:
				system("color C2");
				break;
			case 183:
				system("color C3");
				break;
			case 184:
				system("color C4");
				break;
			case 185:
				system("color C5");
				break;
			case 186:
				system("color C6");
				break;
			case 187:
				system("color C7");
				break;
			case 188:
				system("color C8");
				break;
			case 189:
				system("color C9");
				break;
			case 190:
				system("color CA");
				break;
			case 191:
				system("color CB");
				break;
			case 192:
				system("color CD");
				break;
			case 193:
				system("color CE");
				break;
			case 194:
				system("color CF");
				break;
			case 195:
				system("color D0");
				break;
			case 196:
				system("color D1");
				break;
			case 197:
				system("color D2");
				break;
			case 198:
				system("color D3");
				break;
			case 199:
				system("color D4");
				break;
			case 200:
				system("color D5");
				break;
			case 201:
				system("color D6");
				break;
			case 202:
				system("color D7");
				break;
			case 203:
				system("color D8");
				break;
			case 204:
				system("color D9");
				break;
			case 205:
				system("color DA");
				break;
			case 206:
				system("color DB");
				break;
			case 207:
				system("color DC");
				break;
			case 208:
				system("color DE");
				break;
			case 209:
				system("color DF");
				break;
			case 210:
				system("color E0");
				break;
			case 211:
				system("color E1");
				break;
			case 212:
				system("color E2");
				break;
			case 213:
				system("color E3");
				break;
			case 214:
				system("color E4");
				break;
			case 215:
				system("color E5");
				break;
			case 216:
				system("color E6");
				break;
			case 217:
				system("color E7");
				break;
			case 218:
				system("color E8");
				break;
			case 219:
				system("color E9");
				break;
			case 220:
				system("color EA");
				break;
			case 221:
				system("color EB");
				break;
			case 222:
				system("color EC");
				break;
			case 223:
				system("color ED");
				break;
			case 224:
				system("color EF");
				break;
			case 225:
				system("color F0");
				break;
			case 226:
				system("color F1");
				break;
			case 227:
				system("color F2");
				break;
			case 228:
				system("color F3");
				break;
			case 229:
				system("color F4");
				break;
			case 230:
				system("color F5");
				break;
			case 231:
				system("color F6");
				break;
			case 232:
				system("color F7");
				break;
			case 233:
				system("color F8");
				break;
			case 234:
				system("color F9");
				break;
			case 235:
				system("color FA");
				break;
			case 236:
				system("color FB");
				break;
			case 237:
				system("color FC");
				break;
			case 238:
				system("color FD");
				break;
			case 239:
				system("color FE");
				break;
	}
		return;
}
