using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace zTest
{
    class Program
    {
        static void Main(string[] args)
        {
            //Main123();
            findSmallestNumber();
        }
        static void findSmallestNumber()
        {
            int myInt = 28;
            int i = 1;
            Int64 fact = 1;
            while (true)
            {

                fact =  fact * i ;
                i++;
                if (fact == 0)
                    Console.Write(fact);
                if (myInt == i)
                    break;
            }

            
            int zeroCount = 0;
            int maxZeroCount = 0;
            while (fact > 0)
            {
                if (fact % 10 == 0)
                {
                    zeroCount++;
                }
                else
                {
                    if (maxZeroCount < zeroCount)
                        maxZeroCount = zeroCount;
                    zeroCount = 0;
                }

                fact = fact / 10;
            }


            Console.Write(maxZeroCount);
        }
        public static void Main123()
        {
            int i;
            Int64 fact = 1;
            Int64 number;
            Console.Write("Enter any Number: ");
            number = 12;// int.Parse(Console.ReadLine());
            for (i = 1; i <= number; i++)
            {
                fact = fact * i;
            }
            Console.Write("Factorial of " + number + " is: " + fact);
        }
    }
}
