<h1> Bézier splines </h1>
<h2> Have you ever wondered what smoothing algorithm Excel uses to fit smooth curves on a XY scatter? </h2>
<h2> Have you ever found yourself trying to read Y values off an Excel XY scatter plot? </h2>
<h2> Did you ever wish there was a simple way to linearly interpolate in Excel? </h2>

<h2> If you’ve answered Yes to any of the questions above, you may find this new Excel add-in useful. </h2>


**What does this add-in do?** 
It provides you with 2 new custom Excel functions:

1.	A function to interpolate along Excel’s smooth line fit. This the ‘Bezier’ function.

2.	A function to linearly interpolate and extrapolate. This is the ‘Linerp’ function.

**How *does* Excel compute its smoothing algorithm for curved lines in the XY scatter plots?**
Microsoft is not transparent about this (tisk, tisk). But in fact, it uses a type of parametric curve called a Bézier curve (http://pomax.github.io/bezierinfo/), specifically, a third order Bézier curve with 4 control points.
It is a type of cubic spline and avoids some of the oscillation problems (http://en.wikipedia.org/wiki/Runge's_phenomenon) which typically occur when using high degree polynomials for interpolation.

You can see more of the Bézier function at work in the ‘Examples with Bezier curves.xlsm’ workbook. The tab ‘Testing’ shows how it matches Excel’s smooth line graph for a variety of functional forms.


**How does the Bézier function work, and how do I use it?**
The function allows you to replicate Excel’s smoothing algorithm for curved lines by computing a set of Bézier curves and interpolating the X value on the relevant segment of the spline.

It is very simple to use. In your formula bar, type:
=Bezier(KnownXValues, KnownYValues, XToInterpolate)
where the X and Y values are in columns.

The result is the interpolated value, as if it were on Excel’s curved line.


**If the value you want to interpolate is outside of the range of known X values, you have to option to (linearly) extrapolate backward and forward.**
All you need to do is add the (optional) argument 1, like so:
=Bezier(KnownXValues, KnownYValues, XToInterpolate, 1)

See the Bezier function at work with the parabola below:


**How do I use the linear interpolation/extrapolation function?**
In your formula bar, type:
=Linerp(KnownXValues, KnownYValues, XToInterpolate)

This function can handle data in rows or columns, and (monotonically) increasing or decreasing data.


**How do I get this tool?**

1.	If you want the functions to be available in every Excel workbook.
In Excel, click on File, Options, Add-ins. Scroll down to ‘Manage Excel Add-ins’, and click Go. Click Browse and point to the file ‘Interpolation.xlam’. Make sure you select it, and the functions will be available in any Excel workbook you use.
 

2.	If you just want to use the functions as a one-off, or want to see the underlying VBA code.
In Excel, open your VBA editor by hitting ALT+F11. Insert a Module and paste the code in the text file ‘Bezier and linear interpolation functions.txt’.


**Feedback and testing**
I have tested this code and there are error handlers which should cover most bases. However, if the function(s) bomb, I’d be very grateful if you could let me know.
Suggestions for improvement are most welcome. Currently, the Bézier interpolation function only works for data in columns. If there is any interest, I could rework it so that it accepts data in row vectors.

You are naturally free to modify and redistribute this code.

If you have any questions or feedback, feel free to get in touch at alice.lepissier@gmail.com!

Happy interpolating!
