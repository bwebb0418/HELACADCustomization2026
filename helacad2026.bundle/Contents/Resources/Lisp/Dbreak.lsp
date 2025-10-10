;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;; Programmer's Toolbox - February 1997
;; W.Kramer
;;
;; Dimension line break utility
;;   Break whitness and dimension lines in R13
;;   dimension objects with out exploding the
;;   object.  Demonstrates how dimension objects
;;   are stored and manipulated as well as
;;   dynamic scoping of variables.
;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;; Listing 1:  Main Entry Point
;; Function (C:DBREAK)
;;
(defun C:DBREAK ( /
       ;;Local Bound Variables
       EN  ;;operator selection
       P1  ;;break point 1
       P2  ;;break point 2
       ENL ENB ELB NMB ELT ENT ;;see listing 3
       )
   ;;Operator select input object
   ;;and break point. -see listing 2
   (if (DBREAK0) (progn
      (DBREAK1) ;;See listing 3
      (if (null
         (entmake (list
           '(0 . "BLOCK")
           (assoc 2 ELT)
           (assoc 10 ELT)
           '(70 . 1)
         ))) (progn
          (prompt "\nRedefine failed!")
          (exit) ;;force exit of routine
      )) ;;end IF PROGN
      ;;Loop through block def seeking
      ;;LINE object to change.
      (while ENT
        ;;Get the entity list data
        (setq ELT (entget ENT))
        ;;Does entity match the one
        ;;we want to change?
        (if (not (eq ENL ENT))
           ;;NOT a match, 
           ;;reconstruct block def.
           (entmake ELT)
           ;;else, Found a match,
           ;; do break & reconstruction.
           (DBREAK2) ;;see listing 4
        ) ;;end IF
        ;;get next entity in block def.
        (setq ENT (entnext ENT))
      ) ;;end WHILE
      (if (entmake '((0 . "ENDBLK")))
        (prompt "\nBlock modified")
        (prompt "\nBLOCK END failure!")
      ) ;;end IF
      (entupd ENB) 
     ) ;;end PROGN
     (prompt "\nInvalid input")
   ) ;;end IF dbreak0
   (princ)
)
;;-----------------------------------------------
;; Listing 2 - Function DBREAK0
;;
;; Function: DBREAK0 - user input section
;;
(defun DBREAK0 ( /
       EL ;;entity list of selected object
       )
       ;;Global Variables
       ;; EN entity name of selected object
       ;; P1,P2 points selected
   (setq EN
     (nentsel
       "\nSelect dimension line to break: "))
   ;;
   (if EN (progn
      ;;get the object details
      (setq EL (entget (car EN))
            P1 (cadr EN))
      ;;first check to see if object selected
      ;;was a LINE which was part of an insert.
      (if (and (= (cdr (assoc 0 EL)) "LINE")
               (> (length EN) 2))
          ;;now ask operator to supply the
          ;;other side of the break.
          (setq P2
            (getpoint
              (cadr EN)
                "\nOther side of break: "))
          (prompt "\nLine was not selected!")
      ) ;;end IF
   ))
   ;;Snap selected points to the nearest object
   (if (and P1 P2)
      (setq P1 (osnap (cadr EN) "NEA")
            P2 (osnap P2 "NEA")
      ))
    ;;return True if both input points are okay
   (if (and P1 P2) 'T nil))
;;-----------------------------------------------
;; Listing 3 - Function DBREAK1
;;
;; FUNCTION: DBREAK1 - Set up global variables.
;;
(defun DBREAK1 ()
   ;;Global Variables
   ;;ENL  entity list of line selected
   ;;ENB  entity name of block/dimension
   ;;ELB  entity list of block/dimension
   ;;NMB  name of the block/dimension in table
   ;;ELT  entity list of table entry
   ;;ENT  entity name of block object
   (setq
         ENL (car EN)
         ENB (car (last EN))
         ELB (entget ENB)
         NMB (cdr (assoc 2 ELB))
         ELT (tblsearch "BLOCK" NMB)
         ENT (cdr (assoc -2 ELT))
   )
)
;;-----------------------------------------------
;; Listing 4 - Function DBREAK2
;;
;; Function DBREAK2: convert line object into
;; two line objects at break points.
;;
(defun DBREAK2 ( /
       ;;Local Variables
       LYR   ;;layer name of dim/blk line
       PP1   ;;start point of dim/blk line
       PP2   ;;end point of dim/blk line
       PA    ;;break point 1
       PB    ;;break point 2
       )
       ;;Global Variables
       ;; P1,P2 - operator break points
       ;; ELT - entity list of dim/blk line
   ;;get the line details
   (setq PP1 (cdr (assoc 10 ELT))
         PP2 (cdr (assoc 11 ELT))
         LYR (cdr (assoc 8 ELT)) )
   ;; Determine end points of
   ;;new line segments.
   (if (< (distance PP1 P1)
          (distance PP1 P2))
      (setq PA P1
            PB P2)
      (setq PA P2
            PB P1) )
   ;;Build new lines to replace
   ;;the broken line.
   (LINE PP1 PA Lyr 1)
   (LINE PB PP2 Lyr 1))
;;-----------------------------------------------
;; Listing 5 - Function LINE
;;
;; Function: LINE - create line object
(defun LINE (P1 P2 Lyr BlkFlg)
  (entmake (list
    '(0 . "LINE")
    (cons 6 (if BlkFlg "BYBLOCK" "BYLAYER"))
    (if LYR
      (cons 8 LYR)
      (getvar "CLAYER"))
    (cons 10 P1) (cons 11 P2)
    (cons 62 (if BlkFlg 0 256)))))
;;----------------------------------------------- END OF FILE
