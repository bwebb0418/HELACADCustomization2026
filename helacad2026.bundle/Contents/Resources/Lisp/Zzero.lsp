;
;   ZZERO.LSP
;
;   Wing Tong
;   Designers' Aid
;   666-8th Ave., #1
;   San Francisco, CA 94118
;   415-221-34;
;   RETURNS Z-COORDINATE (THICKNESS) OF ENTITIES TO 0
;
(DEFUN C:ZZERO ()
  (PRINC "\nSELECT OBJECTS [RETURN TO PROCESS ENTIRE DRAWING]: ")
  (SETQ SS1 (SSGET))
;
  (IF (= SS1 NIL)
    (SETQ ENT1 (ENTNEXT))
    (PROGN
      (SETQ N 0)
      (SETQ ENT1 (SSNAME SS1 N))
    )
  )
;
  (WHILE ENT1
    (IF (SETQ PT10 (CDR (ASSOC 10 (ENTGET ENT1))))
      (IF (AND (= (LENGTH PT10) 3) (/= (CAR (REVERSE PT10)) 0))
        (PROGN
          (SETQ NEWPT (CONS 10 (REVERSE (CONS 0.0 (CDR (REVERSE PT10))))))
          (ENTMOD (SUBST NEWPT (ASSOC 10 (ENTGET ENT1)) (ENTGET ENT1)))
        )
      )
    )
;
    (IF (SETQ PT11 (CDR (ASSOC 11 (ENTGET ENT1))))
      (IF (AND (= (LENGTH PT10) 3) (/= (CAR (REVERSE PT11)) 0))
        (PROGN
          (SETQ NEWPT (CONS 11 (REVERSE (CONS 0.0 (CDR (REVERSE PT11))))))
          (ENTMOD (SUBST NEWPT (ASSOC 11 (ENTGET ENT1)) (ENTGET ENT1)))
        )
      )
    )
;
    (IF (= SS1 NIL)
      (SETQ ENT1 (ENTNEXT ENT1))
      (SETQ ENT1 (SSNAME SS1 (SETQ N (1+ N))))
    )
  )
  (PRINC)
)
