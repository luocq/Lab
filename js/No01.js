var Fibonacci=function(len){
		document.write("输出"+len+"位斐波那契数列:</br>");
		for(i=0,j=1,k=0,fib=0;i<len;i++,fib=j+k,j=k,k=fib)
		{
			document.write(fib+"\n");
		}
		document.write("<br/>")		
	}