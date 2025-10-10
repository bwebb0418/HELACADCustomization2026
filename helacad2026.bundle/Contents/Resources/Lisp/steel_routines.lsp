;;;steel Joist
(defun c:sjoist()
	(command "spanlayout" "joist" "steel")
	(princ)
)
;;;steel Beams
(defun c:sbeam()
	(command "-vbarun" "dynamic.steelbeam")
	(princ)
)