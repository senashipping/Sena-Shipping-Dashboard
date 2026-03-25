import React, { useEffect, useState } from "react";
import { useParams, useNavigate } from "react-router-dom";
import SubmissionViewModal from "../../components/SubmissionViewModal";

const SubmissionView: React.FC = () => {
  const { submissionId } = useParams<{ submissionId: string }>();
  const navigate = useNavigate();
  const [isOpen, setIsOpen] = useState(true);

  const handleClose = () => {
    setIsOpen(false);
    // Navigate back to dashboard when modal closes
    navigate("/dashboard");
  };

  useEffect(() => {
    if (!submissionId) {
      navigate("/dashboard");
    }
  }, [submissionId, navigate]);

  if (!submissionId) {
    return null;
  }

  return (
    <SubmissionViewModal
      submissionId={submissionId}
      isOpen={isOpen}
      onClose={handleClose}
    />
  );
};

export default SubmissionView;