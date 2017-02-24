package ru.mishin;

/**
 * Make PedigreeLink
 */
public class PedigreeLink {
    public String toString() {
        String obj = "ParentOrChild: " + this.getParentOrChild()
                + " IndividualId: " + this.getIndividualId()
                + " FamilyId:" + this.getFamilyId();
        return obj;
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
}
