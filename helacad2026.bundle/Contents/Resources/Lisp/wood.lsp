;;;Opens the lumber menu
(defun C:hellumber()
	(vl-vbarun "dynamic.showlumbermenu")
	(princ)
)
;;;Opens the manufactured Wood menu
(defun C:manuwood()
	(vl-vbarun "dynamic.showmanufacturedwoodmenu")
	(princ)
)
;;;Shear Wall
(defun c:shearwall()
	(command "-vbarun" "dynamic.shearwall")
	(princ)
)
;;;Wood post
(defun c:Post()
	(command "-vbarun" "dynamic.post")
	(princ)
)
;;;Wood Joist
(defun c:wjoist()
	(command "spanlayout" "joist" "wood")
	(princ)
)
;;;Wood Beam
(defun c:wbeam()
	(command "-vbarun" "dynamic.wood_beam_tag")
	(princ)
)
;;;Wood Girt
(defun c:wgirt()
	(command "-vbarun" "dynamic.girt_tag")
	(princ)
)
