package ru.mishin;

/**
 * Make PedigreeLink
 */
public class PedigreeLink implements Comparable<PedigreeLink>{
    public String toString() {
        return "ParentOrChild: " + this.getParentOrChild()
                + " IndividualId: " + this.getIndividualId()
                + " FamilyId:" + this.getFamilyId();
    }

    public String getParentOrChild() {
        return parentOrChild;
    }

    public void setParentOrChild(String parentOrChild) {
        this.parentOrChild = parentOrChild;
    }

    public String getFamilyId() {
        return familyId;
    }

    public void setFamilyId(String familyId) {
        this.familyId = familyId;
    }

    public String getIndividualId() {
        return individualId;
    }

    public void setIndividualId(String individualId) {
        this.individualId = individualId;
    }

    private String parentOrChild;
    private String familyId;
    private String individualId;

    @Override
    public int compareTo(PedigreeLink p) {
        int lastCmp = parentOrChild.compareTo(p.parentOrChild);
        if(lastCmp != 0){
            return lastCmp;
        }
        lastCmp = familyId.compareTo(p.familyId);
        return ((lastCmp != 0) ? lastCmp : individualId.compareTo(p.individualId));
    }

    public boolean equals(Object o) {
        if (!(o instanceof PedigreeLink))
            return false;
        PedigreeLink n = (PedigreeLink) o;
        return n.parentOrChild.equals(parentOrChild)
                && n.familyId.equals(familyId)
                && n.individualId.equals(individualId);
    }

}
