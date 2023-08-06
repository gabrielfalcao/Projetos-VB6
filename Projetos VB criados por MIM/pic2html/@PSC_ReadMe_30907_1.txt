Title: Pic2HTML release 2
Description: This is an update from my last version. Someone had suggested using string concatenation and I had completely forgotten to implement that in the first release. I also modified the file saving to be "smarter" so that large files save correctly without taking eons. I was thinking of coding an algorithm to recognize repeating patterns, but I decided it was overkill. This release will do even larger pictures now, so thank you to the person who reminded me of the beauty of using buffers!
 ORIGINAL: 
 I've noticed a lot of "Pic2HTML" programs lately that will turn an image file into an ASCII image of similar colors. While this is a good idea, I figured I'd spend an hour and code an algorithm that was lossless (IE: Final HTML picture is *EXACTLY* the same as the original image file). The algorithm is in the "Generate Code" button and even though it's fairly optimized, it still uses a lot of RAM and takes a while for large image files (it even has some "intelligence" to notice pixel clusters, but that didn't increase the speed: it only decreased the output size). Enjoy.
This file came from Planet-Source-Code.com...the home millions of lines of source code
You can view comments on this code/and or vote on it at: http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=30907&lngWId=1

The author may have retained certain copyrights to this code...please observe their request and the law by reviewing all copyright conditions at the above URL.
