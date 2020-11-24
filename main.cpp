#include<iostream>
#include<random>
using namespace std;
int main(){
	default_random_engine e;
	uniform_real_distribution<double> u(-1.2,3.5);
	cout <<"hello:"<<  u(e) << endl;
	return 0;
}
