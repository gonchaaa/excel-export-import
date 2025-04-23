package com.apache.poi.repository;

import com.apache.poi.entity.User;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import java.time.LocalDate;
import java.util.List;

@Repository
public interface UserRepository extends JpaRepository<User,Long> {
    List<User> findAllByInsertDateBetween(LocalDate startDate, LocalDate endDate);
}
