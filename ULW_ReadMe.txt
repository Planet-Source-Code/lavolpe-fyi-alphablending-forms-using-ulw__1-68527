I have tried several different methods to synch movement of both forms.

1. Using DeferWindowPos doesn't really work. It is more designed for moving child controls within a parent.
The more complex the bkg image (PNG), the more noticable the lag when one form is moved after the other.

2. UpdateLayeredWindows (ULW) allows that smooth per-pixel alphablending of a form based on an image, but it
has one major flaw; controls on the form are invisible, they are painted over by the image.
The only way to hack around this is to subclass each control and reroute its wm_paint event
to an offscreen bitmap then using ULW to update the entire window. This is not only difficult at best,
it is also extremely time consuming and your app can seem sluggish. VISTA may have solved this
problem because its version of ULW allows udpating of only a portion of bitmap/window whereas
the Win2K/XP version always updates the entire window. ULW not only enables per-pixel alpha blending
but it also offers transluceny for the entire window via its Blend parameter or using the 
translucency of the bkg image.  The bkg image must be premultiplied alpha.

3. So if the controls are invisible, what good is ULW? Good question. But since that is out of
our control, how to we work around it? The attached project shows one way and, as always with
workarounds, there are other challenges that must be addressed. To show controls, I created a
separate form that has all the controls on it and the form has a bkg color that won't exist in
any of the controls. Then we use SetLayeredWindowAttributes (SLWA) to make the entire form 
transparent where it contains the bkg color. Then we overlay the controls from onto the bkg form

4. Moving the two together without any noticable lag requires a simple method. Paint the bkg form's
image onto the control form (we lose translucency during the move), hide the bkg form, move the 
controls form. When done, we erase the bkg image from the controls form, show the bkg form and done.

Now all of this is kinda cool, and the attached project doesn't require subclassing. But for more 
advanced techniques and more complicated overlays (if multiple forms are required), then you will
have challenges to overcome. 

I will play with this off and on, but it is only a matter of curiosity for me. I have no immediate
need to solve all the future problems that may be associated with these techniques


LaVolpe