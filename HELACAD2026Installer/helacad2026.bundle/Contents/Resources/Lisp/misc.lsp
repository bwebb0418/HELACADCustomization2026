;;;drain_bearing insert
(defun c:drain_bearing()
	(command "drainbearing" "y")
	(princ)
)
;;;bearing insert
(defun c:bearing()
	(command "drainbearing" "n")
	(princ)
)
;;;Gravel
(defun c:gravel()
	(command "-vbarun" "normal.gravel")
	(princ)
)