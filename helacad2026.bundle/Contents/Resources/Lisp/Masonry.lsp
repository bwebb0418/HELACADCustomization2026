;;;Concrete Block
(defun c:masonry()
	(command "-vbarun" "dynamic.showmasonrymenu")
	(princ)
)
;;;rebar_layout
(defun c:rebar_layout()
	(command "spanlayout" "rebar")
	(princ)
)