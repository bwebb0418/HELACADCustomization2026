;;;View title
(defun c:title()
	(vl-vbarun "title")
	(princ)
)
;;; 2 Part Closed Title
(defun c:2pc_title()
	(command "vtitle" "c")
	(princ)
)
;;; 2 Part Open Title
(defun c:2po_title()
	(command "vtitle" "o")
	(princ)
)
;;; 2 Part Open Section
(defun c:2po_section()
	(command "sect" "o")
	(princ)
)
;;; 2 Part closed Section
(defun c:2pc_section()
	(command "sect" "c")
	(princ)
)
;;; 2 Part Open Detail
(defun c:2po_detail()
	(command "detail" "o")
	(princ)
)
;;; 2 Part closed detail
(defun c:2pc_detail()
	(command "detail" "c")
	(princ)
)