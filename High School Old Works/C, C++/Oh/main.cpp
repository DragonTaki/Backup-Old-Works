#include <iostream>

using namespace std;

void mult(int *a,int *b,bool p)
{
   p?cout<<*a+1<<"*"<<*b+1<<"="<<(*a+1)*(*b+1)<<"\t":cout<<"\t";
   return;
}

int main(void)
{
   const int num=9;
   int numHalf=(num+1)/2;
   for(int j=0;j<numHalf;cout<<endl,j++)
      for(int i=0;i<num;mult(&i,&j,(i>=numHalf-j-1&&i<numHalf+j+1)?((i+(j&1)+(num&1)-1)&1):true),i++);
   for(int j=numHalf;j<num;cout<<endl,j++)
      for(int i=0;i<num;mult(&i,&j,(i>=numHalf-num+j-1&&i<numHalf+num-j-1)?((i+(j&1)+(num&1)-1)&1):true),i++);
   return 0;
}
