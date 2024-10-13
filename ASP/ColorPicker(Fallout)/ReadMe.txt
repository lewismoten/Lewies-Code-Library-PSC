For an odd reason, I just get addicted to playing with interfaces to allow users to choose a color. The logic behind it all is interesting and appears simple. Although figuring it out wasn't such an easy task. 

I've begun to play with a "virtual" drag/drop effect using events such as onMouseDown, onMouseUp to setup a flag, and onMouseMove to change the position of the cursor. I found that moving an element tended to slow things down during the drag/drop process - so I hid the cursor image until the user lets up on the mouse. 

I was finding that draging the mouse was becomming a problem, because as soon as you drag an image outside of the scriptlet, the browser redirects you to that image. The solution was to make all images into background images of SPAN tags. It also appears that the speed of the script executes much faster with this technique. 

I began playing with the scriplets in the main form to talk to each other. For example, when user clicks on a red area, the fallout scriptlet raises the event that it was clicked. The test container sees this event and passes the color information to the luminance scriptlet. the luminance scriptlet changes its color and raises an event signalling that its chosen color has changed. the test container then changes the color of the color sample to the new color.
